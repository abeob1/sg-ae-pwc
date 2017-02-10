Imports System.Windows.Forms
Imports System.Configuration

Namespace AE_PWC_AO06
    Module modMain
        Public p_oApps As SAPbouiCOM.SboGuiApi
        Public WithEvents p_oSBOApplication As SAPbouiCOM.Application
        Public p_oEventHandler As clsEventHandler
        Public p_oDICompany As SAPbobsCOM.Company
        Public p_oTargetCompany As SAPbobsCOM.Company
        Public p_oUICompany As SAPbouiCOM.Company
        Public Const RTN_SUCCESS As Int16 = 1
        Public Const RTN_ERROR As Int16 = 0
        Public ItemListCount As Integer
        Public sItem As String

        Public Const DEBUG_ON As Int16 = 1
        Public Const DEBUG_OFF As Int16 = 0

        Public Const ERR_DISPLAY_STATUS As Int16 = 1
        Public Const ERR_DISPLAY_DIALOGUE As Int16 = 2

        Public p_iDebugMode As Int16
        Public p_iErrDispMethod As Int16
        Public p_iDeleteDebugLog As Int16
        Public sHoldingDB As String

        <STAThread()>
        Sub Main(ByVal args() As String)
            Dim sFuncName As String = "Main"
            Dim sErrDesc As String = String.Empty
            Try
                p_iDebugMode = DEBUG_ON
                p_iErrDispMethod = ERR_DISPLAY_STATUS

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Addon startup function", sFuncName)
                p_oApps = New SAPbouiCOM.SboGuiApi
                'sconn = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056"
                'p_oApps.Connect(args(0))
                p_oApps.Connect(args(0))

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Initializing public SBO Application object", sFuncName)
                p_oSBOApplication = p_oApps.GetApplication

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Retriving SBO application company handle", sFuncName)
                p_oUICompany = p_oSBOApplication.Company


                p_oDICompany = New SAPbobsCOM.Company
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Retrived SBO application company handle", sFuncName)
                ' p_oDICompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
                If Not p_oDICompany.Connected Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectDICompSSO()", sFuncName)
                    If ConnectDICompSSO(p_oDICompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                End If

                'Call WriteToLogFile_Debug("Calling DisplayStatus()", sFuncName)
                'Call DisplayStatus(Nothing, "Addon starting.....please wait....", sErrDesc)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating Event handler class", sFuncName)
                p_oEventHandler = New clsEventHandler(p_oSBOApplication, p_oDICompany)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating FMS", sFuncName)
                AddFMS("AE_ENTITYLIST")
                AddFMS("AE_APRLWINDOW_VENDOR")
                AddFMS("AE_APRLWINDOW_CREATOR")

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating table", sFuncName)
                CreateTable()

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating Procedure", sFuncName)
                CreateProcedure("AE_APPROVALGRID")

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddMenuItems()", sFuncName)
                p_oEventHandler.AddMenuItems()

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling SetApplication Function", sFuncName)
                ' Call p_oEventHandler.SetApplication(sErrDesc)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Addon started successfully", "Main()")

                'Call WriteToLogFile_Debug("Calling EndStatus()", sFuncName)
                ' Call EndStatus(sErrDesc)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Initializing Recordset ", "Main()")

                p_oSBOApplication.StatusBar.SetText("Addon Started Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                System.Windows.Forms.Application.Run()
            Catch ex As Exception

            End Try
        End Sub

        Private Sub CreateTable()
            Dim sFuncName As String = "CreateTable"
            Dim sErrDesc As String = String.Empty
            Try
                If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("HoldingDB")) Then
                    sHoldingDB = ConfigurationManager.AppSettings("HoldingDB")
                End If
                If sHoldingDB = p_oDICompany.CompanyDB Then
                    addField("OUSR", "PASSWORD", "PASSWORD", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
                End If
            Catch ex As Exception
                sErrDesc = ex.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                Throw New ArgumentException(sErrDesc)
            End Try
        End Sub

        Private Sub AddFMS(ByVal sQueryName As String)
            Dim sFuncName As String = "AddFMS"
            Dim sErrDesc As String = String.Empty
            Dim sInternalKey As String = String.Empty
            Dim sSQL As String = String.Empty
            Dim oRecordSet As SAPbobsCOM.Recordset
            Dim sQueryId As String

            Try
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

                oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                sSQL = "SELECT IntrnalKey FROM OUQR WHERE QName = '" + sQueryName + "'"
                oRecordSet.DoQuery(sSQL)
                If oRecordSet.RecordCount > 0 Then
                    sQueryId = oRecordSet.Fields.Item("IntrnalKey").Value

                    If sQueryName = "AE_ENTITYLIST" Then
                        Dim oFMS As SAPbobsCOM.FormattedSearches
                        oFMS = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oFormattedSearches)
                        oFMS.Action = SAPbobsCOM.BoFormattedSearchActionEnum.bofsaQuery
                        oFMS.QueryID = sQueryId
                        oFMS.FormID = "APRL"
                        oFMS.ItemID = "22"
                        oFMS.ColumnID = "-1"

                        oFMS.ByField = SAPbobsCOM.BoYesNoEnum.tNO
                        oFMS.Refresh = SAPbobsCOM.BoYesNoEnum.tNO
                        'oFMS.FieldID = "8"

                        If (oFMS.Add() <> 0) Then
                            sErrDesc = p_oDICompany.GetLastErrorDescription()
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error while adding FMS", sFuncName)
                        End If
                    ElseIf sQueryName = "AE_APRLWINDOW_VENDOR" Then
                        Dim oFMS As SAPbobsCOM.FormattedSearches
                        oFMS = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oFormattedSearches)
                        oFMS.Action = SAPbobsCOM.BoFormattedSearchActionEnum.bofsaQuery
                        oFMS.QueryID = sQueryId
                        oFMS.FormID = "APRL"
                        oFMS.ItemID = "16"
                        oFMS.ColumnID = "-1"

                        oFMS.ByField = SAPbobsCOM.BoYesNoEnum.tNO
                        oFMS.Refresh = SAPbobsCOM.BoYesNoEnum.tNO
                        'oFMS.FieldID = "8"

                        If (oFMS.Add() <> 0) Then
                            sErrDesc = p_oDICompany.GetLastErrorDescription()
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error while adding FMS", sFuncName)
                        End If
                    ElseIf sQueryName = "AE_APRLWINDOW_CREATOR" Then
                        Dim oFMS As SAPbobsCOM.FormattedSearches
                        oFMS = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oFormattedSearches)
                        oFMS.Action = SAPbobsCOM.BoFormattedSearchActionEnum.bofsaQuery
                        oFMS.QueryID = sQueryId
                        oFMS.FormID = "APRL"
                        oFMS.ItemID = "10"
                        oFMS.ColumnID = "-1"

                        oFMS.ByField = SAPbobsCOM.BoYesNoEnum.tNO
                        oFMS.Refresh = SAPbobsCOM.BoYesNoEnum.tNO
                        'oFMS.FieldID = "8"

                        If (oFMS.Add() <> 0) Then
                            sErrDesc = p_oDICompany.GetLastErrorDescription()
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error while adding FMS", sFuncName)
                        End If
                    End If
                Else
                    Dim oQuery As SAPbobsCOM.UserQueries
                    oQuery = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserQueries)
                    If sQueryName = "AE_ENTITYLIST" Then
                        oQuery.Query = "SELECT DISTINCT Name FROM [@AB_COMPANYDATA] "
                        oQuery.QueryCategory = -1
                        oQuery.QueryDescription = sQueryName

                        If oQuery.Add() <> 0 Then
                            sErrDesc = p_oDICompany.GetLastErrorDescription()
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error while adding user query " & sQueryName, sFuncName)
                        Else
                            sInternalKey = p_oDICompany.GetNewObjectKey()

                            sSQL = "SELECT IntrnalKey FROM OUQR WHERE QName = '" + sQueryName + "'"
                            oRecordSet.DoQuery(sSQL)
                            sQueryId = oRecordSet.Fields.Item("IntrnalKey").Value

                            Dim oFMS As SAPbobsCOM.FormattedSearches
                            oFMS = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oFormattedSearches)
                            oFMS.Action = SAPbobsCOM.BoFormattedSearchActionEnum.bofsaQuery
                            oFMS.QueryID = sQueryId
                            oFMS.FormID = "APRL"
                            oFMS.ItemID = "22"
                            oFMS.ColumnID = "-1"

                            oFMS.ByField = SAPbobsCOM.BoYesNoEnum.tNO
                            oFMS.Refresh = SAPbobsCOM.BoYesNoEnum.tNO
                            'oFMS.FieldID = "8"

                            If (oFMS.Add() <> 0) Then
                                sErrDesc = p_oDICompany.GetLastErrorDescription()
                                Call WriteToLogFile(sErrDesc, sFuncName)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error while adding FMS", sFuncName)
                            End If
                        End If
                    ElseIf sQueryName = "AE_APRLWINDOW_VENDOR" Then
                        oQuery.Query = "DECLARE @ENTITY NVARCHAR(MAX) SET @ENTITY = $[$22.uEntity.0] EXEC('SELECT CardName FROM  '+ @ENTITY +'..OCRD WHERE CardType = ''S'' ')"
                        oQuery.QueryCategory = -1
                        oQuery.QueryDescription = sQueryName

                        If oQuery.Add() <> 0 Then
                            sErrDesc = p_oDICompany.GetLastErrorDescription()
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error while adding user query " & sQueryName, sFuncName)
                        Else
                            sSQL = "SELECT IntrnalKey FROM OUQR WHERE QName = '" + sQueryName + "'"
                            oRecordSet.DoQuery(sSQL)
                            sQueryId = oRecordSet.Fields.Item("IntrnalKey").Value

                            Dim oFMS As SAPbobsCOM.FormattedSearches
                            oFMS = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oFormattedSearches)
                            oFMS.Action = SAPbobsCOM.BoFormattedSearchActionEnum.bofsaQuery
                            oFMS.QueryID = sQueryId
                            oFMS.FormID = "APRL"
                            oFMS.ItemID = "16"
                            oFMS.ColumnID = "-1"

                            oFMS.ByField = SAPbobsCOM.BoYesNoEnum.tNO
                            oFMS.Refresh = SAPbobsCOM.BoYesNoEnum.tNO
                            'oFMS.FieldID = "8"

                            If (oFMS.Add() <> 0) Then
                                sErrDesc = p_oDICompany.GetLastErrorDescription()
                                Call WriteToLogFile(sErrDesc, sFuncName)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error while adding FMS", sFuncName)
                            End If
                        End If
                    ElseIf sQueryName = "AE_APRLWINDOW_CREATOR" Then
                        oQuery.Query = "DECLARE @ENTITY NVARCHAR(MAX) SET @ENTITY = $[$22.uEntity.0] EXEC('SELECT DISTINCT U_NAME FROM ' + @ENTITY + '..OWTM T0 INNER JOIN WTM1 T1 ON T0.[WtmCode] = T1.[WtmCode] LEFT JOIN OUSR T2 ON T2.UserID = T1.UserID WHERE T0.Active = ''Y''')"
                        oQuery.QueryCategory = -1
                        oQuery.QueryDescription = sQueryName

                        If oQuery.Add() <> 0 Then
                            sErrDesc = p_oDICompany.GetLastErrorDescription()
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error while adding user query " & sQueryName, sFuncName)
                        Else
                            sSQL = "SELECT IntrnalKey FROM OUQR WHERE QName = '" + sQueryName + "'"
                            oRecordSet.DoQuery(sSQL)
                            sQueryId = oRecordSet.Fields.Item("IntrnalKey").Value

                            Dim oFMS As SAPbobsCOM.FormattedSearches
                            oFMS = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oFormattedSearches)
                            oFMS.Action = SAPbobsCOM.BoFormattedSearchActionEnum.bofsaQuery
                            oFMS.QueryID = sQueryId
                            oFMS.FormID = "APRL"
                            oFMS.ItemID = "10"
                            oFMS.ColumnID = "-1"

                            oFMS.ByField = SAPbobsCOM.BoYesNoEnum.tNO
                            oFMS.Refresh = SAPbobsCOM.BoYesNoEnum.tNO
                            'oFMS.FieldID = "8"

                            If (oFMS.Add() <> 0) Then
                                sErrDesc = p_oDICompany.GetLastErrorDescription()
                                Call WriteToLogFile(sErrDesc, sFuncName)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error while adding FMS", sFuncName)
                            End If
                        End If
                    End If
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with Success", sFuncName)
            Catch ex As Exception
                sErrDesc = ex.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            End Try
        End Sub

        Private Sub CreateProcedure(ByVal sProcedureName As String)
            Dim sFuncName As String = "CreateProcedure"
            Dim sErrDesc As String = String.Empty
            Dim sFile As String = String.Empty
            Dim sTextLine As String = String.Empty
            Dim sSQL As String = String.Empty
            Dim oRecordSet As SAPbobsCOM.Recordset

            Try
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
                sFile = Application.StartupPath & "\" & sProcedureName & ".txt"

                oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                If IO.File.Exists(sFile) Then
                    Using objReader As New System.IO.StreamReader(sFile)
                        sTextLine = objReader.ReadToEnd()
                    End Using
                    If sTextLine <> "" Then
                        sSQL = "SELECT * FROM  sys.procedures where name = '" & sProcedureName & "' "
                        oRecordSet.DoQuery(sSQL)
                        If oRecordSet.RecordCount > 0 Then
                            sSQL = "DROP PROCEDURE " & sProcedureName & " "
                            oRecordSet.DoQuery(sSQL)
                        End If

                        oRecordSet.DoQuery(sTextLine)
                    End If

                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
            Catch ex As Exception
                sErrDesc = ex.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            End Try
        End Sub


    End Module

End Namespace

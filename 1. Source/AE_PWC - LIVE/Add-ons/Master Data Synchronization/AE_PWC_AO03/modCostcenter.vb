
Namespace AE_PWC_AO03



    Module modCostcenter

        Dim sFuncName As String = String.Empty
        Dim ival As Integer
        Dim IsError As Boolean
        Dim iErr As Integer = 0
        Dim sErr As String = String.Empty
        Dim oMatrix As SAPbouiCOM.Matrix = Nothing


        Public Function CostCenterSync(ByRef oHoldingCompany As SAPbobsCOM.Company, ByRef oTragetCompany As SAPbobsCOM.Company, ByVal sMasterdatacode As String, _
                                      ByRef sErrDesc As String) As Long


            'Function   :   CostCenterSync()
            'Purpose    :   
            'Parameters :   ByVal oForm As SAPbouiCOM.Form
            '                   oForm=Form Type
            '               ByRef sErrDesc As String
            '                   sErrDesc=Error Description to be returned to calling function
            '               
            '                   =
            'Return     :   0 - FAILURE
            '               1 - SUCCESS
            'Author     :   SRI
            'Date       :   30/12/2007
            'Change     :

            Dim sFuncName As String = String.Empty
            sFuncName = "CostCenterSync()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function " & sMasterdatacode, sFuncName)
            Dim oRset As SAPbobsCOM.Recordset = oHoldingCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oCmpSrv As SAPbobsCOM.CompanyService = oTragetCompany.GetCompanyService
            Dim oProfitCenterServices As SAPbobsCOM.ProfitCentersService = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.ProfitCentersService)
            Dim oProfitCenter As SAPbobsCOM.ProfitCenter = oProfitCenterServices.GetDataInterface(SAPbobsCOM.ProfitCentersServiceDataInterfaces.pcsProfitCenter)
            Dim oDimension As SAPbobsCOM.IProfitCenterParams = oProfitCenterServices.GetDataInterface(SAPbobsCOM.DimensionsServiceDataInterfaces.dsDimensionParams)
            Dim sSQL As String = String.Empty

            sSQL = "SELECT T0.[PrcCode], T0.[PrcName], T0.[GrpCode], T0.[DimCode], T0.[ValidFrom], T0.[ValidTo], T0.[Active], " & _
               "T0.[U_AB_ENTITY], T0.[U_AB_REPORTCODE], T0.[U_AB_ENTITYNAME], T0.[U_AB_NatCostType], T0.[U_AB_OUCOMMON] FROM OPRC T0 WHERE " & _
               "T0.[PrcCode] = '" & sMasterdatacode & "'"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Cost Center SQL " & sSQL, sFuncName)

            Try

                oRset.DoQuery(sSQL)
                If oRset.RecordCount = 0 Then
                    sErrDesc = "No matching records found in the holding DB " & sMasterdatacode
                    CostCenterSync = RTN_ERROR
                    Call WriteToLogFile("Completed with ERROR ---" & sErrDesc, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErrDesc, sFuncName)
                    Exit Function
                End If

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Assigning informations to Cost Center ", sFuncName)
                oProfitCenter.CenterCode = oRset.Fields.Item("PrcCode").Value
                oProfitCenter.CenterName = oRset.Fields.Item("PrcName").Value
                oProfitCenter.GroupCode = oRset.Fields.Item("GrpCode").Value
                oProfitCenter.InWhichDimension = oRset.Fields.Item("DimCode").Value
                oProfitCenter.Effectivefrom = oRset.Fields.Item("ValidFrom").Value
                oProfitCenter.EffectiveTo = oRset.Fields.Item("ValidTo").Value
                oProfitCenter.Active = SAPbobsCOM.BoYesNoEnum.tYES
                If Not String.IsNullOrEmpty(oRset.Fields.Item("U_AB_ENTITY").Value) Then
                    oProfitCenter.UserFields.Item("U_AB_ENTITY").Value = oRset.Fields.Item("U_AB_ENTITY").Value
                End If

                If Not String.IsNullOrEmpty(oRset.Fields.Item("U_AB_REPORTCODE").Value) Then
                    oProfitCenter.UserFields.Item("U_AB_REPORTCODE").Value = oRset.Fields.Item("U_AB_REPORTCODE").Value
                End If

                If Not String.IsNullOrEmpty(oRset.Fields.Item("U_AB_ENTITYNAME").Value) Then
                    oProfitCenter.UserFields.Item("U_AB_ENTITYNAME").Value = oRset.Fields.Item("U_AB_ENTITYNAME").Value
                End If

                If Not String.IsNullOrEmpty(oRset.Fields.Item("U_AB_NatCostType").Value) Then
                    oProfitCenter.UserFields.Item("U_AB_NatCostType").Value = oRset.Fields.Item("U_AB_NatCostType").Value
                End If

                If Not String.IsNullOrEmpty(oRset.Fields.Item("U_AB_OUCOMMON").Value) Then
                    oProfitCenter.UserFields.Item("U_AB_OUCOMMON").Value = oRset.Fields.Item("U_AB_OUCOMMON").Value
                End If

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add / Update the Item Master Data " & oTragetCompany.CompanyDB, sFuncName)

                Try
                    oProfitCenterServices.AddProfitCenter(oProfitCenter)
                Catch ex As Exception
                    oProfitCenterServices.UpdateProfitCenter(oProfitCenter)
                End Try
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS " & sErrDesc, sFuncName)

                CostCenterSync = RTN_SUCCESS
                sErrDesc = String.Empty
            Catch ex As Exception
                CostCenterSync = RTN_ERROR
                sErrDesc = ex.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            End Try




        End Function








    End Module


End Namespace


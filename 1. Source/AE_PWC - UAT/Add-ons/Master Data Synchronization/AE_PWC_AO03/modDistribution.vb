
Namespace AE_PWC_AO03
    Module modDistribution

        Dim sFuncName As String = String.Empty
        Dim ival As Integer
        Dim IsError As Boolean
        Dim iErr As Integer = 0
        Dim sErr As String = String.Empty
        Dim oMatrix As SAPbouiCOM.Matrix = Nothing




        Public Function DistributionSync(ByRef oHoldingCompany As SAPbobsCOM.Company, ByRef oTragetCompany As SAPbobsCOM.Company, ByVal sMasterdatacode As String, _
                                      ByRef sErrDesc As String) As Long


            'Function   :   DistributionSync()
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


            Dim oRset As SAPbobsCOM.Recordset = Nothing
            Dim oRset_Target As SAPbobsCOM.Recordset = Nothing
            Dim oCmpSrv As SAPbobsCOM.CompanyService = Nothing

            Dim oDistributionRuleServices As SAPbobsCOM.DistributionRulesService = Nothing
            Dim oDistributionRule As SAPbobsCOM.DistributionRule = Nothing
            '  Dim oDimension As SAPbobsCOM.IProfitCenterParams = oProfitCenterServices.GetDataInterface(SAPbobsCOM.DimensionsServiceDataInterfaces.dsDimensionParams)
            Dim oDLParams As SAPbobsCOM.IDistributionRuleParams = Nothing

            Try
                sFuncName = "DistributionSync()"

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function " & sMasterdatacode, sFuncName)
                ' Get distribution rule
                oRset = oHoldingCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRset_Target = oTragetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oCmpSrv = oTragetCompany.GetCompanyService

                oDistributionRuleServices = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.DistributionRulesService)
                oDistributionRule = oDistributionRuleServices.GetDataInterface(SAPbobsCOM.DistributionRulesServiceDataInterfaces.drsDistributionRule)
                oDLParams = oDistributionRuleServices.GetDataInterface(SAPbobsCOM.DistributionRulesServiceDataInterfaces.drsDistributionRuleParams)

                Dim sSQL As String = String.Empty
                Dim sSQL_tar As String = String.Empty
                Dim bAdd As Boolean = False

                sSQL = "SELECT T0.[OcrCode], T0.[OcrName], T0.[OcrTotal], T0.[Direct], T0.[DimCode], T0.[Active], T1.[PrcCode], T1.[PrcAmount], T1.[OcrTotal], " & _
    "T1.[Direct], T1.[ValidFrom], T1.[ValidTo] FROM OOCR T0  INNER JOIN OCR1 T1 ON T0.[OcrCode] = T1.[OcrCode] WHERE T0.[OcrCode] = '" & sMasterdatacode & "'"

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Distribution Rule SQL " & sSQL, sFuncName)

                sSQL_tar = "SELECT T0.[OcrCode], T0.[OcrName] FROM OOCR T0  WHERE T0.[OcrCode] = '" & sMasterdatacode & "'"

                oRset.DoQuery(sSQL)
                If oRset.RecordCount = 0 Then
                    sErrDesc = "No matching records found in the holding DB " & sMasterdatacode
                    DistributionSync = RTN_ERROR
                    Call WriteToLogFile("Completed with ERROR ---" & sErrDesc, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErrDesc, sFuncName)
                    Exit Function
                End If

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Target Distribution Rule SQL " & sSQL_tar, sFuncName)
                oRset_Target.DoQuery(sSQL_tar)
                If oRset_Target.RecordCount = 0 Then
                    bAdd = True
                End If

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Assigning informations to Distribution Rules ", sFuncName)

                oDistributionRule.FactorCode = oRset.Fields.Item("OcrCode").Value
                oDistributionRule.FactorDescription = oRset.Fields.Item("OcrName").Value
                oDistributionRule.InWhichDimension = oRset.Fields.Item("DimCode").Value
                oDistributionRule.TotalFactor = oRset.Fields.Item("OcrTotal").Value

                If oRset.Fields.Item("Active").Value = "Y" Then
                    oDistributionRule.Active = SAPbobsCOM.BoYesNoEnum.tYES
                Else
                    oDistributionRule.Active = SAPbobsCOM.BoYesNoEnum.tNO
                End If

                For imjs As Integer = 0 To oRset.RecordCount - 1
                    oDistributionRule.DistributionRuleLines.Add()
                    oDistributionRule.DistributionRuleLines.Item(imjs).CenterCode = oRset.Fields.Item("PrcCode").Value
                    oDistributionRule.DistributionRuleLines.Item(imjs).TotalInCenter = oRset.Fields.Item("PrcAmount").Value
                    oDistributionRule.DistributionRuleLines.Item(imjs).Effectivefrom = oRset.Fields.Item("ValidFrom").Value
                    oDistributionRule.DistributionRuleLines.Item(imjs).EffectiveTo = oRset.Fields.Item("ValidTo").Value
                    oRset.MoveNext()
                Next imjs

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add / Update the Item Master Data " & oTragetCompany.CompanyDB, sFuncName)
                oDistributionRule.ToXMLFile("Distribution.xml")
                If bAdd = True Then
                    oDistributionRuleServices.AddDistributionRule(oDistributionRule)
                Else
                    oDistributionRuleServices.UpdateDistributionRule(oDistributionRule)
                End If

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS " & sErrDesc, sFuncName)
                DistributionSync = RTN_SUCCESS
                sErrDesc = String.Empty
            Catch ex As Exception
                DistributionSync = RTN_ERROR
                sErrDesc = ex.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            End Try


        End Function

        Public Function DistributionSync_OLD(ByRef oHoldingCompany As SAPbobsCOM.Company, ByRef oTragetCompany As SAPbobsCOM.Company, ByVal sMasterdatacode As String, _
                                      ByRef sErrDesc As String) As Long


            'Function   :   DistributionSync()
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


            Dim oRset As SAPbobsCOM.Recordset = Nothing
            Dim oRset_Target As SAPbobsCOM.Recordset = Nothing
            Dim oCmpSrv As SAPbobsCOM.CompanyService = Nothing

            Dim oDistributionRuleServices As SAPbobsCOM.DistributionRulesService = Nothing
            Dim oDistributionRule As SAPbobsCOM.DistributionRule = Nothing
            '  Dim oDimension As SAPbobsCOM.IProfitCenterParams = oProfitCenterServices.GetDataInterface(SAPbobsCOM.DimensionsServiceDataInterfaces.dsDimensionParams)
            Dim oDLParams As SAPbobsCOM.IDistributionRuleParams = Nothing

            Try
                sFuncName = "DistributionSync()"

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function " & sMasterdatacode, sFuncName)
                ' Get distribution rule
                oRset = oHoldingCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRset_Target = oTragetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oCmpSrv = oTragetCompany.GetCompanyService

                oDistributionRuleServices = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.DistributionRulesService)
                oDistributionRule = oDistributionRuleServices.GetDataInterface(SAPbobsCOM.DistributionRulesServiceDataInterfaces.drsDistributionRule)
                oDLParams = oDistributionRuleServices.GetDataInterface(SAPbobsCOM.DistributionRulesServiceDataInterfaces.drsDistributionRuleParams)

                Dim sSQL As String = String.Empty
                Dim sSQL_tar As String = String.Empty
                Dim bAdd As Boolean = False

                sSQL = "SELECT T0.[OcrCode], T0.[OcrName], T0.[OcrTotal], T0.[Direct], T0.[DimCode], T0.[Active], T1.[PrcCode], T1.[PrcAmount], T1.[OcrTotal], " & _
    "T1.[Direct], T1.[ValidFrom], T1.[ValidTo] FROM OOCR T0  INNER JOIN OCR1 T1 ON T0.[OcrCode] = T1.[OcrCode] WHERE T0.[OcrCode] = '" & sMasterdatacode & "'"

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Distribution Rule SQL " & sSQL, sFuncName)

                sSQL_tar = "SELECT T0.[OcrCode], T0.[OcrName] FROM OOCR T0  WHERE T0.[OcrCode] = '" & sMasterdatacode & "'"

                oRset.DoQuery(sSQL)
                If oRset.RecordCount = 0 Then
                    sErrDesc = "No matching records found in the holding DB " & sMasterdatacode
                    DistributionSync_OLD = RTN_ERROR
                    Call WriteToLogFile("Completed with ERROR ---" & sErrDesc, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErrDesc, sFuncName)
                    Exit Function
                End If

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Target Distribution Rule SQL " & sSQL_tar, sFuncName)
                oRset_Target.DoQuery(sSQL_tar)
                If oRset_Target.RecordCount = 0 Then
                    bAdd = True
                End If

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Assigning informations to Distribution Rules ", sFuncName)

                If bAdd = True Then
                    oDistributionRule.FactorCode = oRset.Fields.Item("OcrCode").Value
                    oDistributionRule.FactorDescription = oRset.Fields.Item("OcrName").Value
                    oDistributionRule.InWhichDimension = oRset.Fields.Item("DimCode").Value
                    oDistributionRule.TotalFactor = oRset.Fields.Item("OcrTotal").Value

                    If oRset.Fields.Item("Active").Value = "Y" Then
                        oDistributionRule.Active = SAPbobsCOM.BoYesNoEnum.tYES
                    Else
                        oDistributionRule.Active = SAPbobsCOM.BoYesNoEnum.tNO
                    End If
                    oDistributionRule.Direct = oRset.Fields.Item("Direct").Value

                    For imjs As Integer = 0 To oRset.RecordCount - 1
                        oDistributionRule.DistributionRuleLines.Add()
                        oDistributionRule.DistributionRuleLines.Item(imjs).CenterCode = oRset.Fields.Item("PrcCode").Value
                        oDistributionRule.DistributionRuleLines.Item(imjs).TotalInCenter = oRset.Fields.Item("PrcAmount").Value
                        oDistributionRule.DistributionRuleLines.Item(imjs).Effectivefrom = oRset.Fields.Item("ValidFrom").Value
                        oDistributionRule.DistributionRuleLines.Item(imjs).EffectiveTo = oRset.Fields.Item("ValidTo").Value

                        oRset.MoveNext()
                    Next imjs

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add the Distribution Master Data " & oTragetCompany.CompanyDB, sFuncName)
                    oDistributionRuleServices.AddDistributionRule(oDistributionRule)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding successfully", sFuncName)

                Else
                    oDLParams.FactorCode = oRset.Fields.Item("OcrCode").Value
                    oDistributionRule = oDistributionRuleServices.GetDistributionRule(oDLParams)
                    oDistributionRule.FactorDescription = oRset.Fields.Item("OcrName").Value
                    oDistributionRule.InWhichDimension = oRset.Fields.Item("DimCode").Value
                    oDistributionRule.TotalFactor = oRset.Fields.Item("OcrTotal").Value
                    If oRset.Fields.Item("Active").Value = "Y" Then
                        oDistributionRule.Active = SAPbobsCOM.BoYesNoEnum.tYES
                    Else
                        oDistributionRule.Active = SAPbobsCOM.BoYesNoEnum.tNO
                    End If
                    oDistributionRule.Direct = oRset.Fields.Item("Direct").Value
                    For imjs As Integer = 0 To oRset.RecordCount - 1

                        oDistributionRule.DistributionRuleLines.Add()
                        oDistributionRule.DistributionRuleLines.Item(imjs).CenterCode = oRset.Fields.Item("PrcCode").Value
                        oDistributionRule.DistributionRuleLines.Item(imjs).TotalInCenter = oRset.Fields.Item("PrcAmount").Value
                        oDistributionRule.DistributionRuleLines.Item(imjs).Effectivefrom = oRset.Fields.Item("ValidFrom").Value
                        oDistributionRule.DistributionRuleLines.Item(imjs).EffectiveTo = oRset.Fields.Item("ValidTo").Value
                        oRset.MoveNext()
                    Next imjs
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Update the Distribution Master Data " & oTragetCompany.CompanyDB, sFuncName)
                    oDistributionRuleServices.UpdateDistributionRule(oDistributionRule)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Updated successfully", sFuncName)
                End If
                oDistributionRule.ToXMLFile("Distribution.xml")
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS " & sErrDesc, sFuncName)
                DistributionSync_OLD = RTN_SUCCESS
                sErrDesc = String.Empty
            Catch ex As Exception
                DistributionSync_OLD = RTN_ERROR
                sErrDesc = ex.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            End Try


        End Function


    End Module
End Namespace


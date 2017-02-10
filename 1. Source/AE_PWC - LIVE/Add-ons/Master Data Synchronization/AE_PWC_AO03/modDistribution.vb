
Namespace AE_PWC_AO03
    Module modDistribution

        Dim sFuncName As String = String.Empty
        Dim ival As Integer
        Dim IsError As Boolean
        Dim iErr As Integer = 0
        Dim sErr As String = String.Empty
        Dim oMatrix As SAPbouiCOM.Matrix = Nothing

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
            sFuncName = "DistributionSync()"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function " & sMasterdatacode, sFuncName)

            Dim oRset As SAPbobsCOM.Recordset = oHoldingCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oCmpSrv As SAPbobsCOM.CompanyService = oTragetCompany.GetCompanyService

            Dim oDistributionRuleServices As SAPbobsCOM.DistributionRulesService = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.DistributionRulesService)
            Dim oDistributionRule As SAPbobsCOM.DistributionRule = oDistributionRuleServices.GetDataInterface(SAPbobsCOM.DistributionRulesServiceDataInterfaces.drsDistributionRule)
            '  Dim oDimension As SAPbobsCOM.IProfitCenterParams = oProfitCenterServices.GetDataInterface(SAPbobsCOM.DimensionsServiceDataInterfaces.dsDimensionParams)
            Dim sSQL As String = String.Empty

            sSQL = "SELECT T0.[OcrCode], T0.[OcrName], T0.[OcrTotal], T0.[Direct], T0.[DimCode], T0.[Active], T1.[PrcCode], T1.[PrcAmount], T1.[OcrTotal], " & _
"T1.[Direct], T1.[ValidFrom], T1.[ValidTo] FROM OOCR T0  INNER JOIN OCR1 T1 ON T0.[OcrCode] = T1.[OcrCode] WHERE T0.[OcrCode] = '" & sMasterdatacode & "'"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Distribution Rule SQL " & sSQL, sFuncName)

            Try

                oRset.DoQuery(sSQL)
                If oRset.RecordCount = 0 Then
                    sErrDesc = "No matching records found in the holding DB " & sMasterdatacode
                    DistributionSync_OLD = RTN_ERROR
                    Call WriteToLogFile("Completed with ERROR ---" & sErrDesc, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErrDesc, sFuncName)
                    Exit Function
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

                Try
                    ' oDistributionRule.ToXMLFile("C:\distrbution.xml")
                    oDistributionRuleServices.AddDistributionRule(oDistributionRule)

                Catch ex As Exception
                    oDistributionRuleServices.UpdateDistributionRule(oDistributionRule)
                End Try
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

        Public Function DistributionSync_1(ByRef oHoldingCompany As SAPbobsCOM.Company, ByRef oTragetCompany As SAPbobsCOM.Company, ByVal sMasterdatacode As String, _
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
            Dim oCmpSrvT As SAPbobsCOM.CompanyService = Nothing
            Dim oDVDR As DataView = Nothing
            Dim oDistributionRuleServices As SAPbobsCOM.DistributionRulesService = Nothing
            Dim oDistributionRule As SAPbobsCOM.DistributionRule = Nothing
            Dim oDistributionRuleServicesT As SAPbobsCOM.DistributionRulesService = Nothing
            Dim oDistributionRuleT As SAPbobsCOM.DistributionRule = Nothing
            '  Dim oDimension As SAPbobsCOM.IProfitCenterParams = oProfitCenterServices.GetDataInterface(SAPbobsCOM.DimensionsServiceDataInterfaces.dsDimensionParams)
            Dim oDLParams As SAPbobsCOM.IDistributionRuleParams = Nothing
            Dim imjs As Integer = 0
            Dim icount As Integer = 0
            Dim dAmount As Double = 0

            Try
                sFuncName = "DistributionSync()"

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function " & sMasterdatacode, sFuncName)
                ' Get distribution rule
                oRset = oHoldingCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRset_Target = oTragetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oCmpSrv = oTragetCompany.GetCompanyService
                oCmpSrvT = oTragetCompany.GetCompanyService
                Dim oDLService As SAPbobsCOM.DistributionRulesService = oCmpSrvT.GetBusinessService(SAPbobsCOM.ServiceTypes.DistributionRulesService)
                oDistributionRuleServices = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.DistributionRulesService)
                oDistributionRule = oDistributionRuleServices.GetDataInterface(SAPbobsCOM.DistributionRulesServiceDataInterfaces.drsDistributionRule)
                oDLParams = oDistributionRuleServices.GetDataInterface(SAPbobsCOM.DistributionRulesServiceDataInterfaces.drsDistributionRuleParams)
                oDistributionRuleServicesT = oCmpSrvT.GetBusinessService(SAPbobsCOM.ServiceTypes.DistributionRulesService)
                oDistributionRuleT = oDistributionRuleServicesT.GetDataInterface(SAPbobsCOM.DistributionRulesServiceDataInterfaces.drsDistributionRule)
                Dim sSQL As String = String.Empty
                Dim sSQL_tar As String = String.Empty
                Dim bAdd As Boolean = False

                sSQL = "SELECT T0.[OcrCode], T0.[OcrName], T0.[OcrTotal], T0.[Direct], T0.[DimCode], T0.[Active], T1.[PrcCode], T1.[PrcAmount], " & _
    "T1.[ValidFrom], T1.[ValidTo], " & _
    "isnull((SELECT case when  TT1.[PrcCode]  = '' then 'N' else 'E' end from " & oTragetCompany.CompanyDB & " ..OOCR TT0 INNER JOIN " & oTragetCompany.CompanyDB & " ..OCR1 TT1 ON TT0.[OcrCode] = TT1.[OcrCode]  " & _
    "where TT1.[OcrCode] = T0.[OcrCode] and TT1.[PrcCode] = T1.[PrcCode] and TT1.[ValidFrom] = T1.[ValidFrom]) ,'N') [E/N]" & _
    " FROM OOCR T0  INNER JOIN OCR1 T1 ON T0.[OcrCode] = T1.[OcrCode] WHERE T0.[OcrCode] = '" & sMasterdatacode & "'"

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Distribution Rule SQL " & sSQL, sFuncName)

                sSQL_tar = "SELECT T0.[OcrCode], T0.[OcrName] FROM OOCR T0  WHERE T0.[OcrCode] = '" & sMasterdatacode & "'"

                oRset.DoQuery(sSQL)
                If oRset.RecordCount = 0 Then
                    sErrDesc = "No matching records found in the holding DB " & sMasterdatacode
                    DistributionSync_1 = RTN_ERROR
                    Call WriteToLogFile("Completed with ERROR ---" & sErrDesc, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErrDesc, sFuncName)
                    Exit Function
                End If

                oDVDR = New DataView(ConvertRecordset(oRset, sErrDesc))
                If sErrDesc.Length > 0 Then Throw New ArgumentException(sErrDesc)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Target Distribution Rule SQL " & sSQL_tar, sFuncName)
                oRset_Target.DoQuery(sSQL_tar)
                If oRset_Target.RecordCount = 0 Then
                    bAdd = True
                End If

                If bAdd = True Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Assigning informations to Distribution Rules - Add() ", sFuncName)

                    oDistributionRule.FactorCode = oRset.Fields.Item("OcrCode").Value
                    oDistributionRule.FactorDescription = oRset.Fields.Item("OcrName").Value
                    oDistributionRule.InWhichDimension = oRset.Fields.Item("DimCode").Value
                    oDistributionRule.TotalFactor = oRset.Fields.Item("OcrTotal").Value

                    If oRset.Fields.Item("Active").Value = "Y" Then
                        oDistributionRule.Active = SAPbobsCOM.BoYesNoEnum.tYES
                    Else
                        oDistributionRule.Active = SAPbobsCOM.BoYesNoEnum.tNO
                    End If


                    For imjs = 0 To oRset.RecordCount - 1
                        oDistributionRule.DistributionRuleLines.Add()
                        oDistributionRule.DistributionRuleLines.Item(imjs).CenterCode = oRset.Fields.Item("PrcCode").Value
                        oDistributionRule.DistributionRuleLines.Item(imjs).TotalInCenter = oRset.Fields.Item("PrcAmount").Value
                        oDistributionRule.DistributionRuleLines.Item(imjs).Effectivefrom = oRset.Fields.Item("ValidFrom").Value
                        oDistributionRule.DistributionRuleLines.Item(imjs).EffectiveTo = oRset.Fields.Item("ValidTo").Value
                        oRset.MoveNext()
                    Next imjs
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add " & oTragetCompany.CompanyDB, sFuncName)
                    oDistributionRule.ToXMLFile("Distribution.xml")
                    oDistributionRuleServices.AddDistributionRule(oDistributionRule)
                Else
                    imjs = 0
                    oDVDR.RowFilter = "[E/N]='E'"
                    If oDVDR.Count > 0 Then
                        icount = oDVDR.Count
                        dAmount = 0
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting Existing Line " & oTragetCompany.CompanyDB, sFuncName)
                        oDistributionRule.FactorCode = oDVDR.Item(0)("OcrCode").ToString
                        oDistributionRule.FactorDescription = oDVDR.Item(0)("OcrName").ToString
                        oDistributionRule.InWhichDimension = oDVDR.Item(0)("DimCode").ToString

                        If oDVDR.Item(0)("Active").ToString = "Y" Then
                            oDistributionRule.Active = SAPbobsCOM.BoYesNoEnum.tYES
                        Else
                            oDistributionRule.Active = SAPbobsCOM.BoYesNoEnum.tNO
                        End If
                        For Each odrv As DataRowView In oDVDR
                            oDistributionRule.DistributionRuleLines.Add()
                            oDistributionRule.DistributionRuleLines.Item(imjs).CenterCode = odrv("PrcCode").ToString
                            oDistributionRule.DistributionRuleLines.Item(imjs).TotalInCenter = odrv("PrcAmount").ToString
                            dAmount += odrv("PrcAmount").ToString
                            oDistributionRule.DistributionRuleLines.Item(imjs).Effectivefrom = odrv("ValidFrom").ToString
                            oDistributionRule.DistributionRuleLines.Item(imjs).EffectiveTo = odrv("ValidTo").ToString
                            imjs += 1
                        Next
                        oDistributionRule.TotalFactor = dAmount ''oDVDR.Item(0)("OcrTotal").ToString
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Update - Existing " & oTragetCompany.CompanyDB, sFuncName)
                        oDistributionRule.ToXMLFile("Distribution.xml")
                        oDistributionRuleServices.UpdateDistributionRule(oDistributionRule)
                    End If
                    imjs = oDVDR.Count
                    oDVDR.RowFilter = "[E/N]='N'"

                    If oDVDR.Count > 0 Then
                        dAmount = 0
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting New Line " & oTragetCompany.CompanyDB, sFuncName)
                        oDLParams.FactorCode = oDVDR.Item(0)("OcrCode").ToString
                        Dim oDL As SAPbobsCOM.DistributionRule = oDLService.GetDistributionRule(oDLParams)
                        oDL.FactorDescription = oDVDR.Item(0)("OcrName").ToString
                        oDL.InWhichDimension = oDVDR.Item(0)("DimCode").ToString

                        If oDVDR.Item(0)("Active").ToString = "Y" Then
                            oDistributionRule.Active = SAPbobsCOM.BoYesNoEnum.tYES
                        Else
                            oDistributionRule.Active = SAPbobsCOM.BoYesNoEnum.tNO
                        End If
                        For Each odrv As DataRowView In oDVDR
                            oDL.DistributionRuleLines.Add()
                            oDL.DistributionRuleLines.Item(imjs).CenterCode = odrv("PrcCode").ToString
                            oDL.DistributionRuleLines.Item(imjs).TotalInCenter = odrv("PrcAmount").ToString
                            dAmount += odrv("PrcAmount").ToString
                            oDL.DistributionRuleLines.Item(imjs).Effectivefrom = odrv("ValidFrom").ToString
                            oDL.DistributionRuleLines.Item(imjs).EffectiveTo = odrv("ValidTo").ToString
                            imjs += 1
                        Next
                        oDL.TotalFactor = dAmount
                        oDLService.UpdateDistributionRule(oDL)
                    End If
                End If
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS " & sErrDesc, sFuncName)
                DistributionSync_1 = RTN_SUCCESS
                sErrDesc = String.Empty
            Catch ex As Exception
                DistributionSync_1 = RTN_ERROR
                sErrDesc = ex.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDistributionRuleServices)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDistributionRule)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDistributionRuleServicesT)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDistributionRuleT)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDLParams)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRset)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRset_Target)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCmpSrv)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCmpSrvT)
                oDistributionRuleServices = Nothing
                oDistributionRule = Nothing
                oDistributionRuleServicesT = Nothing
                oDistributionRuleT = Nothing
                oDLParams = Nothing
                oRset_Target = Nothing
                oCmpSrv = Nothing
                oCmpSrvT = Nothing
            End Try
        End Function

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
            Dim oRset_N As SAPbobsCOM.Recordset = Nothing
            Dim oRset_Target As SAPbobsCOM.Recordset = Nothing
            Dim oCmpSrv As SAPbobsCOM.CompanyService = Nothing
            Dim oCmpSrvT As SAPbobsCOM.CompanyService = Nothing
            Dim oDVDR As DataView = Nothing
            Dim oDVDR_1 As DataView = Nothing
            Dim oDistributionRuleServices As SAPbobsCOM.DistributionRulesService = Nothing
            Dim oDistributionRule As SAPbobsCOM.DistributionRule = Nothing
            Dim oDistributionRuleServicesT As SAPbobsCOM.DistributionRulesService = Nothing
            Dim oDistributionRuleT As SAPbobsCOM.DistributionRule = Nothing
            '  Dim oDimension As SAPbobsCOM.IProfitCenterParams = oProfitCenterServices.GetDataInterface(SAPbobsCOM.DimensionsServiceDataInterfaces.dsDimensionParams)
            Dim oDLParams As SAPbobsCOM.IDistributionRuleParams = Nothing
            Dim imjs As Integer = 0
            Dim icount As Integer = 0
            Dim iTGCount As Integer = 0
            Dim dAmount As Double = 0
            Dim oDTDistinct As DataTable = Nothing
            Dim oDTDistinct_1 As DataTable = Nothing

            Try
                sFuncName = "DistributionSync()"

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function " & sMasterdatacode, sFuncName)
                ' Get distribution rule
                oRset = oHoldingCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRset_Target = oTragetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRset_N = oTragetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oCmpSrv = oTragetCompany.GetCompanyService
                oCmpSrvT = oTragetCompany.GetCompanyService
                Dim oDLService As SAPbobsCOM.DistributionRulesService = oCmpSrvT.GetBusinessService(SAPbobsCOM.ServiceTypes.DistributionRulesService)
                oDistributionRuleServices = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.DistributionRulesService)
                oDistributionRule = oDistributionRuleServices.GetDataInterface(SAPbobsCOM.DistributionRulesServiceDataInterfaces.drsDistributionRule)
                oDLParams = oDistributionRuleServices.GetDataInterface(SAPbobsCOM.DistributionRulesServiceDataInterfaces.drsDistributionRuleParams)
                oDistributionRuleServicesT = oCmpSrvT.GetBusinessService(SAPbobsCOM.ServiceTypes.DistributionRulesService)
                oDistributionRuleT = oDistributionRuleServicesT.GetDataInterface(SAPbobsCOM.DistributionRulesServiceDataInterfaces.drsDistributionRule)
                Dim sSQL As String = String.Empty
                Dim sSQL_tar As String = String.Empty
                Dim sSQL_N As String = String.Empty
                Dim bAdd As Boolean = False

                sErrDesc = String.Empty
                ''Query Commented on 28 Sep 2016
                '              sSQL = "SELECT T0.[OcrCode], T0.[OcrName], T0.[OcrTotal], isnull(T0.[Direct],'N') [Direct] , T0.[DimCode], isnull(T0.[Active],'N') [Active], T1.[PrcCode], T1.[PrcAmount], " & _
                '  "T1.[ValidFrom], T1.[ValidTo], " & _
                '  "isnull((SELECT case when  TT1.[PrcCode]  = '' then 'N' else 'E' end from " & oTragetCompany.CompanyDB & " ..OOCR TT0 INNER JOIN " & oTragetCompany.CompanyDB & " ..OCR1 TT1 ON TT0.[OcrCode] = TT1.[OcrCode]  " & _
                '  "where TT1.[OcrCode] = T0.[OcrCode] and TT1.[PrcCode] = T1.[PrcCode] and TT1.[ValidFrom] = T1.[ValidFrom]) ,'N') [E/N] ,  T1.[OcrTotal] [LineTotal]," & _
                '  " isnull((SELECT count(TT0.OcrCode ) from SEAC ..OOCR TT0 INNER JOIN SEAC ..OCR1 TT1 ON TT0.[OcrCode] = TT1.[OcrCode]  " & _
                ' "where TT1.[OcrCode] = T0.[OcrCode] and TT0.OcrCode = T0.OcrCode and TT1.PrcCode <> '') ,0) [count], " & _
                ' "(SELECT count(TT1.OcrCode ) " & _
                '"from SEAC ..OOCR TT0 INNER JOIN SEAC ..OCR1 TT1 ON TT0.[OcrCode] = TT1.[OcrCode]  " & _
                '"where TT1.[OcrCode] = T0.[OcrCode] and TT1.[ValidFrom] = T1.[ValidFrom])  [NCount]" & _
                '  " FROM OOCR T0  INNER JOIN OCR1 T1 ON T0.[OcrCode] = T1.[OcrCode] WHERE T0.[OcrCode] = '" & sMasterdatacode & "'"

                ''Query Changed on 28 Sep 2016
                sSQL = "SELECT T0.[OcrCode], T0.[OcrName], T0.[OcrTotal], isnull(T0.[Direct],'N') [Direct] , T0.[DimCode], isnull(T0.[Active],'N') [Active], T1.[PrcCode], T1.[PrcAmount], " & _
                        "T1.[ValidFrom], T1.[ValidTo], " & _
                        "isnull((SELECT case when  TT1.[PrcCode]  = '' then 'N' else 'E' end from " & oTragetCompany.CompanyDB & " ..OOCR TT0 INNER JOIN " & oTragetCompany.CompanyDB & " ..OCR1 TT1 ON TT0.[OcrCode] = TT1.[OcrCode]  " & _
                        "where TT1.[OcrCode] = T0.[OcrCode] and TT1.[PrcCode] = T1.[PrcCode] and TT1.[ValidFrom] = T1.[ValidFrom]) ,'N') [E/N] ,  T1.[OcrTotal] [LineTotal]," & _
                        " isnull((SELECT count(TT0.OcrCode ) from  " & oTragetCompany.CompanyDB & " ..OOCR TT0 INNER JOIN  " & oTragetCompany.CompanyDB & " ..OCR1 TT1 ON TT0.[OcrCode] = TT1.[OcrCode]  " & _
                        "where TT1.[OcrCode] = T0.[OcrCode] and TT0.OcrCode = T0.OcrCode and TT1.PrcCode <> '') ,0) [count], " & _
                        "(SELECT count(TT1.OcrCode ) " & _
                        "from  " & oTragetCompany.CompanyDB & " ..OOCR TT0 INNER JOIN  " & oTragetCompany.CompanyDB & " ..OCR1 TT1 ON TT0.[OcrCode] = TT1.[OcrCode]  " & _
                        "where TT1.[OcrCode] = T0.[OcrCode] and TT1.[ValidFrom] = T1.[ValidFrom])  [NCount]" & _
                        " ,isnull((SELECT count(TT0.OcrCode ) from " & oTragetCompany.CompanyDB & " ..OOCR TT0 INNER JOIN " & oTragetCompany.CompanyDB & " ..OCR1 TT1 " & _
                        "ON TT0.[OcrCode] = TT1.[OcrCode]  where TT1.[OcrCode] = T0.[OcrCode] and TT0.OcrCode = T0.OcrCode and TT1.PrcCode <> '') ,0) [TGcount]" & _
                        " FROM OOCR T0  INNER JOIN OCR1 T1 ON T0.[OcrCode] = T1.[OcrCode] WHERE T0.[OcrCode] = '" & sMasterdatacode & "'"



                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Distribution Rule SQL " & sSQL, sFuncName)

                sSQL_tar = "SELECT T0.[OcrCode], T0.[OcrName] FROM OOCR T0  WHERE T0.[OcrCode] = '" & sMasterdatacode & "'"

                '' oTragetCompany.StartTransaction()
                oRset.DoQuery(sSQL)
                If oRset.RecordCount = 0 Then
                    sErrDesc = "No matching records found in the holding DB " & sMasterdatacode
                    DistributionSync = RTN_ERROR
                    Call WriteToLogFile("Completed with ERROR ---" & sErrDesc, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErrDesc, sFuncName)
                    Exit Function
                End If

                oDVDR = New DataView(ConvertRecordset(oRset, sErrDesc))
                If sErrDesc.Length > 0 Then Throw New ArgumentException(sErrDesc)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Target Distribution Rule SQL " & sSQL_tar, sFuncName)
                oRset_Target.DoQuery(sSQL_tar)
                If oRset_Target.RecordCount = 0 Then
                    bAdd = True
                End If

                If bAdd = True Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Assigning informations to Distribution Rules - Add() ", sFuncName)

                    If oDVDR.Count > 0 Then
                        oDTDistinct = New DataTable
                        oDTDistinct = oDVDR.ToTable
                        oDTDistinct = oDTDistinct.DefaultView.ToTable(True, "ValidFrom")
                        icount = oDVDR.Count
                        dAmount = 0
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting Existing Line " & oTragetCompany.CompanyDB, sFuncName)
                        '--- New Distribution rules 
                        For Each odtrow As DataRow In oDTDistinct.Rows
                            '' oDVDR.RowFilter = "ValidFrom='" & odtrow(0).ToString() & "' and [E/N]='E'"
                            oDVDR.RowFilter = "ValidFrom='" & odtrow(0).ToString() & "'"
                            If oDVDR.Count > 0 Then
                                If DistributionParameterAssign_ADD(oDVDR, oTragetCompany, imjs, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                imjs = imjs
                            End If
                        Next
                    End If

                Else
                    ''Commented my john 03Oct2016 
                    ''imjs = 0
                    ''oDVDR.RowFilter = "NCount>0 and [E/N]='E'"
                    ''If oDVDR.Count > 0 Then
                    ''    oDTDistinct = New DataTable
                    ''    oDTDistinct = oDVDR.ToTable
                    ''    oDTDistinct = oDTDistinct.DefaultView.ToTable(True, "ValidFrom")
                    ''    icount = oDVDR.Count
                    ''    dAmount = 0
                    ''    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting Existing Line " & oTragetCompany.CompanyDB, sFuncName)

                    ''    For Each odtrow As DataRow In oDTDistinct.Rows
                    ''        '' oDVDR.RowFilter = "ValidFrom='" & odtrow(0).ToString() & "' and [E/N]='E'"
                    ''        oDVDR.RowFilter = "ValidFrom='" & odtrow(0).ToString() & "' and NCount>0 and [E/N]='E'"
                    ''        If oDVDR.Count > 0 Then
                    ''            ''Threading.Thread.Sleep(6000)
                    ''            If DistributionParameterAssign_(oDVDR, oTragetCompany, imjs, sErrDesc) <> RTN_SUCCESS Then
                    ''                Call WriteToLogFile(sErrDesc, sFuncName)
                    ''                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sMasterdatacode & " " & sErrDesc, sFuncName)
                    ''            End If
                    ''            '' Throw New ArgumentException(sErrDesc)
                    ''            imjs = imjs
                    ''        End If
                    ''    Next
                    ''End If

                    sSQL_N = "SELECT isnull(count(TT0.OcrCode ),0) [TGcount] from " & oTragetCompany.CompanyDB & " ..OOCR TT0 INNER JOIN " & oTragetCompany.CompanyDB & " ..OCR1 TT1 " & _
                       "ON TT0.[OcrCode] = TT1.[OcrCode]  where TT1.[OcrCode] = '" & sMasterdatacode & "'and TT1.PrcCode <> ''"

                    ''sSQL_N = "SELECT T0.[OcrCode], T0.[OcrName], T0.[OcrTotal], isnull(T0.[Direct],'N') [Direct] , T0.[DimCode], isnull(T0.[Active],'N') [Active], T1.[PrcCode], T1.[PrcAmount], " & _
                    ''   "T1.[ValidFrom], T1.[ValidTo], " & _
                    ''   "isnull((SELECT case when  TT1.[PrcCode]  = '' then 'N' else 'E' end from " & oTragetCompany.CompanyDB & " ..OOCR TT0 INNER JOIN " & oTragetCompany.CompanyDB & " ..OCR1 TT1 ON TT0.[OcrCode] = TT1.[OcrCode]  " & _
                    ''   "where TT1.[OcrCode] = T0.[OcrCode] and TT1.[PrcCode] = T1.[PrcCode] and TT1.[ValidFrom] = T1.[ValidFrom]) ,'N') [E/N] ,  T1.[OcrTotal] [LineTotal]," & _
                    ''   " isnull((SELECT count(TT0.OcrCode ) from  " & oTragetCompany.CompanyDB & " ..OOCR TT0 INNER JOIN  " & oTragetCompany.CompanyDB & " ..OCR1 TT1 ON TT0.[OcrCode] = TT1.[OcrCode]  " & _
                    ''   "where TT1.[OcrCode] = T0.[OcrCode] and TT0.OcrCode = T0.OcrCode and TT1.PrcCode <> '') ,0) [count], " & _
                    ''   "(SELECT count(TT1.OcrCode ) " & _
                    ''   "from  " & oTragetCompany.CompanyDB & " ..OOCR TT0 INNER JOIN  " & oTragetCompany.CompanyDB & " ..OCR1 TT1 ON TT0.[OcrCode] = TT1.[OcrCode]  " & _
                    ''   "where TT1.[OcrCode] = T0.[OcrCode] and TT1.[ValidFrom] = T1.[ValidFrom])  [NCount]" & _
                    ''   " ,isnull((SELECT count(TT0.OcrCode ) from " & oTragetCompany.CompanyDB & " ..OOCR TT0 INNER JOIN " & oTragetCompany.CompanyDB & " ..OCR1 TT1 " & _
                    ''   "ON TT0.[OcrCode] = TT1.[OcrCode]  where TT1.[OcrCode] = T0.[OcrCode] and TT0.OcrCode = T0.OcrCode and TT1.PrcCode <> '') ,0) [TGcount]" & _
                    ''   " FROM OOCR T0  INNER JOIN OCR1 T1 ON T0.[OcrCode] = T1.[OcrCode] WHERE T0.[OcrCode] = '" & sMasterdatacode & "'"



                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Distribution Rule SQL - To get TGCount " & sSQL_N, sFuncName)
                    oRset_N.DoQuery(sSQL_N)
                    If oRset_N.RecordCount = 0 Then
                        sErrDesc = "No matching records found in the holding DB " & sMasterdatacode
                        DistributionSync = RTN_ERROR
                        Call WriteToLogFile("Completed with ERROR ---" & sErrDesc, sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErrDesc, sFuncName)
                        Exit Function
                    End If
                    If oRset_N.RecordCount > 0 Then
                        iTGCount = oRset_N.Fields.Item("TGCount").Value
                    End If

                    oDVDR.RowFilter = "[E/N]='N'"

                    If oDVDR.Count > 0 Then
                        dAmount = 0
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting New Line " & oTragetCompany.CompanyDB, sFuncName)
                        oDTDistinct = New DataTable
                        oDTDistinct = oDVDR.ToTable
                        oDTDistinct = oDTDistinct.DefaultView.ToTable(True, "ValidFrom")
                        ''Code commented - 28Sep 2016
                        'imjs = oDVDR.Item(0)("count")
                        'imjs = oDVDR.Item(0)("TGcount")
                        imjs = iTGCount
                        '' imjs = 0                   
                        For Each odtrow As DataRow In oDTDistinct.Rows
                            oDVDR.RowFilter = "ValidFrom='" & odtrow(0).ToString() & "' and [E/N]='N'"
                            Threading.Thread.Sleep(8000)
                            If DistributionParameterAssign(oDVDR, oTragetCompany, imjs, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            ''If oDVDR.Count > 0 Then
                            ''    oDVDR_1 = oDVDR
                            ''    oDTDistinct_1 = oDVDR_1.ToTable
                            ''    oDTDistinct_1 = oDTDistinct_1.DefaultView.ToTable(True, "PrcCode")
                            ''    For Each odt As DataRow In oDTDistinct_1.Rows
                            ''        oDVDR_1.RowFilter = "PrcCode='" & odt(0).ToString() & "' and [E/N]='N'"
                            ''        If DistributionParameterAssign(oDVDR_1, oTragetCompany, imjs, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            ''    Next
                            ''End If
                        Next
                    End If
                End If
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS " & sErrDesc, sFuncName)
                DistributionSync = RTN_SUCCESS
                sErrDesc = String.Empty
            Catch ex As Exception
                DistributionSync = RTN_ERROR
                sErrDesc = ex.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDistributionRuleServices)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDistributionRule)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDistributionRuleServicesT)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDistributionRuleT)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDLParams)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRset)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRset_Target)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCmpSrv)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCmpSrvT)
                oDistributionRuleServices = Nothing
                oDistributionRule = Nothing
                oDistributionRuleServicesT = Nothing
                oDistributionRuleT = Nothing
                oDLParams = Nothing
                oRset_Target = Nothing
                oCmpSrv = Nothing
                oCmpSrvT = Nothing
            End Try
        End Function

        Public Function DistributionParameterAssign(ByVal oDVDR As DataView, ByRef oTragetCompany As SAPbobsCOM.Company, ByRef imjs As Integer, ByRef sErrDesc As String) As Long

            Dim oDistributionRuleServicesT As SAPbobsCOM.DistributionRulesService = Nothing
            Dim oDistributionRuleT As SAPbobsCOM.DistributionRule = Nothing
            Dim oDL As SAPbobsCOM.DistributionRule = Nothing
            '  Dim oDimension As SAPbobsCOM.IProfitCenterParams = oProfitCenterServices.GetDataInterface(SAPbobsCOM.DimensionsServiceDataInterfaces.dsDimensionParams)
            Dim oDLParams As SAPbobsCOM.IDistributionRuleParams = Nothing
            Dim oCmpSrvT As SAPbobsCOM.CompanyService = Nothing
            oCmpSrvT = oTragetCompany.GetCompanyService
            Dim oDLService As SAPbobsCOM.DistributionRulesService = oCmpSrvT.GetBusinessService(SAPbobsCOM.ServiceTypes.DistributionRulesService)
            oDistributionRuleServicesT = oCmpSrvT.GetBusinessService(SAPbobsCOM.ServiceTypes.DistributionRulesService)
            oDistributionRuleT = oDistributionRuleServicesT.GetDataInterface(SAPbobsCOM.DistributionRulesServiceDataInterfaces.drsDistributionRule)
            oDLParams = oDistributionRuleServicesT.GetDataInterface(SAPbobsCOM.DistributionRulesServiceDataInterfaces.drsDistributionRuleParams)
            ''  oDL = oDistributionRuleServicesT.GetDataInterface(SAPbobsCOM.DistributionRulesServiceDataInterfaces.drsDistributionRule)
            Dim damount As Double = 0.0
            Dim sFuncName = String.Empty

            Try
                sFuncName = "ItemMaterSync()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

                ''Commented on 27Sep2016
                oDLParams.FactorCode = oDVDR.Item(0)("OcrCode").ToString
                oDL = oDLService.GetDistributionRule(oDLParams)

                Dim sStirng1 As String = oDL.ToXMLString()
                '' oDL.FactorCode = oDVDR.Item(0)("OcrCode").ToString
                oDL.FactorDescription = oDVDR.Item(0)("OcrName").ToString
                oDL.InWhichDimension = oDVDR.Item(0)("DimCode").ToString

                If oDVDR.Item(0)("Active").ToString = "Y" Then
                    oDL.Active = SAPbobsCOM.BoYesNoEnum.tYES
                Else
                    oDL.Active = SAPbobsCOM.BoYesNoEnum.tNO
                End If

                oDL.Direct = oDVDR.Item(0)("Direct").ToString

                For Each odrv As DataRowView In oDVDR
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Loop " & odrv("PrcCode").ToString, sFuncName)
                    oDL.DistributionRuleLines.Add()
                    oDL.DistributionRuleLines.Item(imjs).CenterCode = odrv("PrcCode").ToString
                    damount += odrv("PrcAmount").ToString
                    oDL.DistributionRuleLines.Item(imjs).TotalInCenter = odrv("PrcAmount").ToString
                    oDL.DistributionRuleLines.Item(imjs).Effectivefrom = odrv("ValidFrom").ToString
                    oDL.DistributionRuleLines.Item(imjs).EffectiveTo = odrv("ValidTo").ToString
                    imjs += 1
                Next

                Dim dTotalfactor As Double = oDVDR.Item(0)("LineTotal")
                oDL.TotalFactor = dTotalfactor + 0.001

                Dim sStirng As String = oDL.ToXMLString()
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("XML " & sStirng, sFuncName)
                oDLService.UpdateDistributionRule(oDL)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS " & sErrDesc, sFuncName)
                DistributionParameterAssign = RTN_SUCCESS
                sErrDesc = String.Empty
            Catch ex As Exception
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with Error " & sErrDesc, sFuncName)
                DistributionParameterAssign = RTN_ERROR
                sErrDesc = ex.Message
            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDistributionRuleServicesT)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDistributionRuleT)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDLParams)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCmpSrvT)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDL)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDLService)
                oDistributionRuleServicesT = Nothing
                oDistributionRuleT = Nothing
                oDLParams = Nothing
                oCmpSrvT = Nothing
                oDL = Nothing
                oDLService = Nothing
            End Try


        End Function

        Public Function DistributionParameterAssign_(ByVal oDVDR As DataView, ByRef oTragetCompany As SAPbobsCOM.Company, ByVal imjs As Integer, ByRef sErrDesc As String) As Long

            Dim oDistributionRuleServices As SAPbobsCOM.DistributionRulesService = Nothing
            Dim oDistributionRule As SAPbobsCOM.DistributionRule = Nothing

            '  Dim oDimension As SAPbobsCOM.IProfitCenterParams = oProfitCenterServices.GetDataInterface(SAPbobsCOM.DimensionsServiceDataInterfaces.dsDimensionParams)
            Dim oDLParams As SAPbobsCOM.IDistributionRuleParams = Nothing
            Dim oCmpSrvT As SAPbobsCOM.CompanyService = Nothing
            oCmpSrvT = oTragetCompany.GetCompanyService
            Dim oDLService As SAPbobsCOM.DistributionRulesService = oCmpSrvT.GetBusinessService(SAPbobsCOM.ServiceTypes.DistributionRulesService)
            oDistributionRuleServices = oCmpSrvT.GetBusinessService(SAPbobsCOM.ServiceTypes.DistributionRulesService)
            oDistributionRule = oDistributionRuleServices.GetDataInterface(SAPbobsCOM.DistributionRulesServiceDataInterfaces.drsDistributionRule)
            oDLParams = oDistributionRuleServices.GetDataInterface(SAPbobsCOM.DistributionRulesServiceDataInterfaces.drsDistributionRuleParams)

            Try


                oDistributionRule.FactorCode = oDVDR.Item(0)("OcrCode").ToString
                oDistributionRule.FactorDescription = oDVDR.Item(0)("OcrName").ToString
                oDistributionRule.InWhichDimension = oDVDR.Item(0)("DimCode").ToString

                If oDVDR.Item(0)("Active").ToString = "Y" Then
                    oDistributionRule.Active = SAPbobsCOM.BoYesNoEnum.tYES
                Else
                    oDistributionRule.Active = SAPbobsCOM.BoYesNoEnum.tNO
                End If

                oDistributionRule.Direct = oDVDR.Item(0)("Direct").ToString

                For Each odrv As DataRowView In oDVDR
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(odrv("PrcCode").ToString, sFuncName)
                    oDistributionRule.DistributionRuleLines.Add()
                    oDistributionRule.DistributionRuleLines.Item(imjs).CenterCode = odrv("PrcCode").ToString
                    oDistributionRule.DistributionRuleLines.Item(imjs).TotalInCenter = odrv("PrcAmount").ToString
                    '' dAmount += odrv("PrcAmount").ToString
                    oDistributionRule.DistributionRuleLines.Item(imjs).Effectivefrom = odrv("ValidFrom").ToString
                    If odrv("ValidTo").ToString <> "30/12/1899 12:00:00 AM" Then
                        oDistributionRule.DistributionRuleLines.Item(imjs).EffectiveTo = odrv("ValidTo").ToString
                    End If
                    imjs += 1
                Next
                oDistributionRule.TotalFactor = oDVDR.Item(0)("LineTotal").ToString
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Update - Existing " & oTragetCompany.CompanyDB, sFuncName)
                oDistributionRule.ToXMLFile("Distribution.xml")
                Dim sXML As String = oDistributionRule.ToXMLString()

                oDistributionRuleServices.UpdateDistributionRule(oDistributionRule)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS " & sErrDesc, sFuncName)
                DistributionParameterAssign_ = RTN_SUCCESS

                sErrDesc = String.Empty
            Catch ex As Exception
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with Error " & sErrDesc, sFuncName)
                DistributionParameterAssign_ = RTN_ERROR
                sErrDesc = ex.Message
            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDistributionRuleServices)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDistributionRule)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDLParams)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCmpSrvT)
                oDistributionRuleServices = Nothing
                oDistributionRule = Nothing
                oDLParams = Nothing
                oCmpSrvT = Nothing
            End Try


        End Function

        Public Function DistributionParameterAssign_ADD(ByVal oDVDR As DataView, ByRef oTragetCompany As SAPbobsCOM.Company, ByVal imjs As Integer, ByRef sErrDesc As String) As Long

            Dim oDistributionRuleServices As SAPbobsCOM.DistributionRulesService = Nothing
            Dim oDistributionRule As SAPbobsCOM.DistributionRule = Nothing

            '  Dim oDimension As SAPbobsCOM.IProfitCenterParams = oProfitCenterServices.GetDataInterface(SAPbobsCOM.DimensionsServiceDataInterfaces.dsDimensionParams)
            Dim oDLParams As SAPbobsCOM.IDistributionRuleParams = Nothing
            Dim oCmpSrvT As SAPbobsCOM.CompanyService = Nothing
            oCmpSrvT = oTragetCompany.GetCompanyService
            Dim oDLService As SAPbobsCOM.DistributionRulesService = oCmpSrvT.GetBusinessService(SAPbobsCOM.ServiceTypes.DistributionRulesService)
            oDistributionRuleServices = oCmpSrvT.GetBusinessService(SAPbobsCOM.ServiceTypes.DistributionRulesService)
            oDistributionRule = oDistributionRuleServices.GetDataInterface(SAPbobsCOM.DistributionRulesServiceDataInterfaces.drsDistributionRule)
            oDLParams = oDistributionRuleServices.GetDataInterface(SAPbobsCOM.DistributionRulesServiceDataInterfaces.drsDistributionRuleParams)

            Try
                Dim dAmount As Double = 0.0
                oDistributionRule.FactorCode = oDVDR.Item(0)("OcrCode").ToString
                oDistributionRule.FactorDescription = oDVDR.Item(0)("OcrName").ToString
                oDistributionRule.InWhichDimension = oDVDR.Item(0)("DimCode").ToString

                If oDVDR.Item(0)("Active").ToString = "Y" Then
                    oDistributionRule.Active = SAPbobsCOM.BoYesNoEnum.tYES
                Else
                    oDistributionRule.Active = SAPbobsCOM.BoYesNoEnum.tNO
                End If

                oDistributionRule.Direct = oDVDR.Item(0)("Direct").ToString

                For Each odrv As DataRowView In oDVDR
                    oDistributionRule.DistributionRuleLines.Add()
                    oDistributionRule.DistributionRuleLines.Item(imjs).CenterCode = odrv("PrcCode").ToString
                    oDistributionRule.DistributionRuleLines.Item(imjs).TotalInCenter = odrv("PrcAmount").ToString
                    dAmount += odrv("PrcAmount").ToString
                    oDistributionRule.DistributionRuleLines.Item(imjs).Effectivefrom = odrv("ValidFrom").ToString
                    '' oDistributionRule.DistributionRuleLines.Item(imjs).EffectiveTo = odrv("ValidTo").ToString
                    imjs += 1
                Next
                oDistributionRule.TotalFactor = dAmount + 0.001 ''oDVDR.Item(0)("LineTotal").ToString
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to ADD " & oTragetCompany.CompanyDB, sFuncName)
                oDistributionRule.ToXMLFile("Distribution.xml")
                oDistributionRuleServices.AddDistributionRule(oDistributionRule)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS " & sErrDesc, sFuncName)
                DistributionParameterAssign_ADD = RTN_SUCCESS

                sErrDesc = String.Empty
            Catch ex As Exception
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with Error " & sErrDesc, sFuncName)
                DistributionParameterAssign_ADD = RTN_ERROR
                sErrDesc = ex.Message
            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDistributionRuleServices)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDistributionRule)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDLParams)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCmpSrvT)
                oDistributionRuleServices = Nothing
                oDistributionRule = Nothing
                oDLParams = Nothing
                oCmpSrvT = Nothing
            End Try


        End Function
    End Module
End Namespace


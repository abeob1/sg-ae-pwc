
Imports System.Xml

Namespace AE_PWC_AO03

    Module modChartofAccounts

        Dim sPath As String = String.Empty
        Dim sFuncName As String = String.Empty
        Dim ival As Integer
        Dim IsError As Boolean
        Dim iErr As Integer = 0
        Dim sErr As String = String.Empty
        Dim xDoc As New XmlDocument
        Dim oMatrix As SAPbouiCOM.Matrix = Nothing


        Public Function ChartofAccounts(ByRef oHoldingCompany As SAPbobsCOM.Company, ByRef oTragetCompany As SAPbobsCOM.Company, ByVal sMasterdatacode As String, _
                                        ByRef sErrDesc As String) As Long

            'Function   :   ChartofAccounts()
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

            Dim sFuncName As String
            Dim sSqlquery As String = String.Empty
            Dim oChartofAccounts As SAPbobsCOM.ChartOfAccounts = Nothing
            Dim oChartofAccount As SAPbobsCOM.ChartOfAccounts = Nothing
            Dim oRest_Holding As SAPbobsCOM.Recordset = Nothing
            Dim oRest_Target As SAPbobsCOM.Recordset = Nothing
            Dim iGrpLine As Integer = 0
            Try
                sFuncName = "ChartofAccounts()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function " & sMasterdatacode, sFuncName)

                oChartofAccounts = oHoldingCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oChartOfAccounts)
                oChartofAccount = oTragetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oChartOfAccounts)
                oRest_Holding = oHoldingCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRest_Target = oTragetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting COA Sync Function" & oTragetCompany.CompanyDB, sFuncName)


                If oChartofAccounts.GetByKey(sMasterdatacode) Then
                    If oChartofAccount.GetByKey(sMasterdatacode) Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Caling Function COA_Assignment()", sFuncName)
                        COA_Assignment(oChartofAccount, oChartofAccounts)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add / Update the COA " & oTragetCompany.CompanyDB, sFuncName)
                        ival = oChartofAccount.Update()
                        If ival <> 0 Then
                            IsError = True
                            oTragetCompany.GetLastError(iErr, sErr)
                            Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
                            sErrDesc = sErr
                            ChartofAccounts = RTN_ERROR
                            Exit Function
                                End If
                    Else
                        oChartofAccount.Code = oChartofAccounts.Code
                        COA_Assignment(oChartofAccount, oChartofAccounts)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add / Update the COA " & oTragetCompany.CompanyDB, sFuncName)
                        ival = oChartofAccount.Add()
                        If ival <> 0 Then
                            IsError = True
                            oTragetCompany.GetLastError(iErr, sErr)
                            Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
                            sErrDesc = sErr
                            ChartofAccounts = RTN_ERROR
                            Exit Function

                        End If
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS " & sErrDesc, sFuncName)
                Else
                    sErrDesc = "No matching records found in the holding DB " & sMasterdatacode
                    ChartofAccounts = RTN_ERROR
                    Call WriteToLogFile("Completed with ERROR ---" & sErrDesc, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErrDesc, sFuncName)
                    Exit Function
                    ' If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS " & oTragetCompany.CompanyDB, sFuncName)
                End If

                ChartofAccounts = RTN_SUCCESS
                sErrDesc = String.Empty
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

            Catch exc As Exception
                ChartofAccounts = RTN_ERROR
                sErrDesc = exc.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Finally

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oChartofAccounts)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oChartofAccount)
                oRest_Holding = Nothing
                oRest_Target = Nothing

            End Try
        End Function

        Public Sub COA_Assignment(ByRef oChartofAccount As SAPbobsCOM.ChartOfAccounts, ByRef oChartofAccounts As SAPbobsCOM.ChartOfAccounts)

            sFuncName = "COA_Assignment()"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

            oChartofAccount.Name = oChartofAccounts.Name
            oChartofAccount.AcctCurrency = oChartofAccounts.AcctCurrency
            'oChartofAccount.TaxExemptAccount = SAPbobsCOM.BoYesNoEnum.tYES
            'oChartofAccount.TaxLiableAccount = SAPbobsCOM.BoYesNoEnum.tYES
            'oChartofAccount.RevaluationCoordinated = SAPbobsCOM.BoYesNoEnum.tYES
            ''oChartofAccount.ReconciledAccount = SAPbobsCOM.BoYesNoEnum.tYES
            ''oChartofAccount.LoadingType = SAPbobsCOM.BoYesNoEnum.tYES
            ''oChartofAccount.LiableForAdvances = SAPbobsCOM.BoYesNoEnum.tYES
            ''oChartofAccount.DatevFirstDataEntry = SAPbobsCOM.BoYesNoEnum.tYES
            ''oChartofAccount.DatevAutoAccount = SAPbobsCOM.BoYesNoEnum.tYES
            ''oChartofAccount.CashAccount = SAPbobsCOM.BoYesNoEnum.tYES
            ''oChartofAccount.BudgetAccount = SAPbobsCOM.BoYesNoEnum.tYES


            If oChartofAccounts.AccountType = SAPbobsCOM.BoAccountTypes.at_Expenses Then
                oChartofAccount.AccountType = SAPbobsCOM.BoAccountTypes.at_Expenses
            ElseIf oChartofAccounts.AccountType = SAPbobsCOM.BoAccountTypes.at_Other Then
                oChartofAccount.AccountType = SAPbobsCOM.BoAccountTypes.at_Other
            ElseIf oChartofAccounts.AccountType = SAPbobsCOM.BoAccountTypes.at_Revenues Then
                oChartofAccount.AccountType = SAPbobsCOM.BoAccountTypes.at_Revenues
            End If

            If oChartofAccounts.LockManualTransaction = SAPbobsCOM.BoYesNoEnum.tYES Then
                oChartofAccount.LockManualTransaction = SAPbobsCOM.BoYesNoEnum.tYES
            Else
                oChartofAccount.LockManualTransaction = SAPbobsCOM.BoYesNoEnum.tNO
            End If

            If oChartofAccounts.CashAccount = SAPbobsCOM.BoYesNoEnum.tYES Then
                oChartofAccount.CashAccount = SAPbobsCOM.BoYesNoEnum.tYES
            Else
                oChartofAccount.CashAccount = SAPbobsCOM.BoYesNoEnum.tNO
            End If

            If oChartofAccounts.DistributionRuleRelevant = SAPbobsCOM.BoYesNoEnum.tYES Then
                oChartofAccount.DistributionRuleRelevant = SAPbobsCOM.BoYesNoEnum.tYES
            Else
                oChartofAccount.DistributionRuleRelevant = SAPbobsCOM.BoYesNoEnum.tNO
            End If

            If oChartofAccounts.DistributionRule2Relevant = SAPbobsCOM.BoYesNoEnum.tYES Then
                oChartofAccount.DistributionRule2Relevant = SAPbobsCOM.BoYesNoEnum.tYES

            Else
                oChartofAccount.DistributionRule2Relevant = SAPbobsCOM.BoYesNoEnum.tNO
            End If
            If oChartofAccounts.DistributionRule3Relevant = SAPbobsCOM.BoYesNoEnum.tYES Then
                oChartofAccount.DistributionRule3Relevant = SAPbobsCOM.BoYesNoEnum.tYES
            Else
                oChartofAccount.DistributionRule3Relevant = SAPbobsCOM.BoYesNoEnum.tNO
            End If
            If oChartofAccounts.DistributionRule4Relevant = SAPbobsCOM.BoYesNoEnum.tYES Then
                oChartofAccount.DistributionRule4Relevant = SAPbobsCOM.BoYesNoEnum.tYES
            Else
                oChartofAccount.DistributionRule4Relevant = SAPbobsCOM.BoYesNoEnum.tNO
            End If
            oChartofAccount.DataExportCode = oChartofAccounts.DataExportCode

            If oChartofAccounts.ActiveAccount = SAPbobsCOM.BoYesNoEnum.tYES Then
                oChartofAccount.ActiveAccount = SAPbobsCOM.BoYesNoEnum.tYES
            Else
                oChartofAccount.ActiveAccount = SAPbobsCOM.BoYesNoEnum.tNO
            End If

            oChartofAccount.FrozenFor = SAPbobsCOM.BoYesNoEnum.tNO

            ''If oChartofAccounts.FrozenFor = SAPbobsCOM.BoYesNoEnum.tYES Then
            ''    oChartofAccount.FrozenFor = SAPbobsCOM.BoYesNoEnum.tYES
            ''Else
            ''    oChartofAccount.FrozenFor = SAPbobsCOM.BoYesNoEnum.tNO
            ''End If
            oChartofAccount.FrozenFrom = oChartofAccounts.FrozenFrom
            oChartofAccount.FrozenTo = oChartofAccounts.FrozenTo

            If oChartofAccounts.BudgetAccount = SAPbobsCOM.BoYesNoEnum.tYES Then
                oChartofAccount.BudgetAccount = SAPbobsCOM.BoYesNoEnum.tYES
            Else
                oChartofAccount.BudgetAccount = SAPbobsCOM.BoYesNoEnum.tNO
            End If
            If oChartofAccounts.AllowChangeVatGroup = SAPbobsCOM.BoYesNoEnum.tYES Then
                oChartofAccount.AllowChangeVatGroup = SAPbobsCOM.BoYesNoEnum.tYES
            Else
                oChartofAccount.AllowChangeVatGroup = SAPbobsCOM.BoYesNoEnum.tNO
            End If
            oChartofAccount.FatherAccountKey = oChartofAccounts.FatherAccountKey

            
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS ", sFuncName)

        End Sub

    End Module
End Namespace





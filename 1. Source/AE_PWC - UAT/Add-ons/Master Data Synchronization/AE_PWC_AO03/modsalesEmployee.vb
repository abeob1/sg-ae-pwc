


Namespace AE_PWC_AO03

    Module modsalesEmployee
        Dim sFuncName As String = String.Empty
        Dim ival As Integer
        Dim IsError As Boolean
        Dim iErr As Integer = 0
        Dim sErr As String = String.Empty
        Dim oMatrix As SAPbouiCOM.Matrix = Nothing
        Dim sSQL As String = String.Empty


        Public Function SalesEmployeeSync(ByRef oHoldingCompany As SAPbobsCOM.Company, ByRef oTragetCompany As SAPbobsCOM.Company, ByVal sMasterdatacode As String, _
                                   ByRef sErrDesc As String) As Long

            'Function   :   UsersSync()
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
            sFuncName = "SalesEmployeeSync()"
            Dim sSalesEmployee As Integer

            Try

                Dim oSalesEmployee_Holding As SAPbobsCOM.SalesPersons = oHoldingCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oSalesPersons)
                Dim oSalesEmployee_Target As SAPbobsCOM.SalesPersons = oTragetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oSalesPersons)

                Dim oRset As SAPbobsCOM.Recordset = oHoldingCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim oCmpSrv As SAPbobsCOM.CompanyService = oTragetCompany.GetCompanyService
                Dim oDepartmentServicesT As SAPbobsCOM.DepartmentsService = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.DepartmentsService)
                Dim oDepartmentT As SAPbobsCOM.Department = oDepartmentServicesT.GetDataInterface(SAPbobsCOM.DepartmentsServiceDataInterfaces.dsDepartment)

                Dim oCmpSrvH As SAPbobsCOM.CompanyService = oHoldingCompany.GetCompanyService
                Dim oDepartmentServicesH As SAPbobsCOM.DepartmentsService = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.DepartmentsService)
                Dim oDepartmentH As SAPbobsCOM.Department = oDepartmentServicesH.GetDataInterface(SAPbobsCOM.DepartmentsServiceDataInterfaces.dsDepartment)

                '  Dim sFileName As String = sPath & "\SalesPerson.xml"
                oRset.DoQuery("SELECT T0.[SlpCode], T0.[SlpName], T0.[Memo] FROM OSLP T0 WHERE T0.[SlpName] = '" & sMasterdatacode & "'")
                sSalesEmployee = oRset.Fields.Item("SlpCode").Value


                If oSalesEmployee_Holding.GetByKey(sSalesEmployee) Then

                    Department(oDepartmentH, oDepartmentServicesH, oSalesEmployee_Holding)

                    ''oDepartmentH.Name = Left(Trim(oSalesEmployee_Holding.SalesEmployeeName), 20)
                    ''oDepartmentH.Description = oSalesEmployee_Holding.SalesEmployeeName
                    ''Try
                    ''    oDepartmentServicesH.AddDepartment(oDepartmentH)
                    ''Catch ex As Exception
                    ''    oDepartmentServicesH.UpdateDepartment(oDepartmentH)
                    ''End Try

                    If oSalesEmployee_Target.GetByKey(sSalesEmployee) Then

                        SalesEmployee(oSalesEmployee_Target, oSalesEmployee_Holding)
                        ''oSalesEmployee_Target.SalesEmployeeName = oSalesEmployee_Holding.SalesEmployeeName
                        ''oSalesEmployee_Target.Remarks = oSalesEmployee_Holding.Remarks
                        ''If oSalesEmployee_Holding.Active = SAPbobsCOM.BoYesNoEnum.tYES Then
                        ''    oSalesEmployee_Target.Active = SAPbobsCOM.BoYesNoEnum.tYES
                        ''Else
                        ''    oSalesEmployee_Target.Active = SAPbobsCOM.BoYesNoEnum.tNO
                        ''End If
                        ival = oSalesEmployee_Target.Update()
                        If ival <> 0 Then
                            IsError = True
                            oTragetCompany.GetLastError(iErr, sErr)
                            Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
                            sErrDesc = sErr
                            SalesEmployeeSync = RTN_ERROR
                            Exit Function
                        End If

                        Department(oDepartmentT, oDepartmentServicesT, oSalesEmployee_Holding)
                        ''oDepartmentT.Name = Left(Trim(oSalesEmployee_Holding.SalesEmployeeName), 20)
                        ''oDepartmentT.Description = oSalesEmployee_Holding.SalesEmployeeName
                        ''If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add / Update the Item Master Data " & oTragetCompany.CompanyDB, sFuncName)

                        ''Try
                        ''    oDepartmentServicesT.AddDepartment(oDepartmentT)

                        ''Catch ex As Exception
                        ''    oDepartmentServicesT.UpdateDepartment(oDepartmentT)
                        ''End Try
                    Else
                        SalesEmployee(oSalesEmployee_Target, oSalesEmployee_Holding)
                        ''oSalesEmployee_Target.SalesEmployeeName = oSalesEmployee_Holding.SalesEmployeeName
                        ''oSalesEmployee_Target.Remarks = oSalesEmployee_Holding.Remarks
                        ''If oSalesEmployee_Holding.Active = SAPbobsCOM.BoYesNoEnum.tYES Then
                        ''    oSalesEmployee_Target.Active = SAPbobsCOM.BoYesNoEnum.tYES
                        ''Else
                        ''    oSalesEmployee_Target.Active = SAPbobsCOM.BoYesNoEnum.tNO
                        ''End If
                        ival = oSalesEmployee_Target.Add()
                        If ival <> 0 Then
                            IsError = True
                            oTragetCompany.GetLastError(iErr, sErr)
                            Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
                            sErrDesc = sErr
                            SalesEmployeeSync = RTN_ERROR
                            Exit Function
                        End If

                        Department(oDepartmentT, oDepartmentServicesT, oSalesEmployee_Holding)
                        ' ''oDepartmentT.Name = oSalesEmployee_Holding.SalesEmployeeName
                        ' ''oDepartmentT.Description = oSalesEmployee_Holding.SalesEmployeeName
                        ' ''If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add / Update the Item Master Data " & oTragetCompany.CompanyDB, sFuncName)

                        ' ''Try
                        ' ''    oDepartmentServicesT.AddDistributionRule(oDepartmentT)

                        ' ''Catch ex As Exception
                        ' ''    oDepartmentServicesT.UpdateDistributionRule(oDepartmentT)
                        ' ''End Try
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS " & sErrDesc, sFuncName)
                Else
                    sSQL = "SELECT T0.[Name], T0.[Remarks], T0.[Father] FROM OUDP T0 WHERE T0.[Name] = '" & sMasterdatacode & "'"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(" Department SQL ", sFuncName)
                    oRset.DoQuery(sSQL)
                    If oRset.RecordCount > 0 Then

                        If oSalesEmployee_Holding.GetByKey(sSalesEmployee) Then
                            SalesEmployee(oSalesEmployee_Holding, oRset.Fields.Item("Name").Value, oRset.Fields.Item("Remarks").Value)
                            ival = oSalesEmployee_Target.Update()
                            If ival <> 0 Then
                                IsError = True
                                oTragetCompany.GetLastError(iErr, sErr)
                                Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
                                sErrDesc = sErr
                                SalesEmployeeSync = RTN_ERROR
                                Exit Function
                            End If

                        Else
                            SalesEmployee(oSalesEmployee_Holding, oRset.Fields.Item("Name").Value, oRset.Fields.Item("Remarks").Value)
                            ival = oSalesEmployee_Target.Add()
                            If ival <> 0 Then
                                IsError = True
                                oTragetCompany.GetLastError(iErr, sErr)
                                Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
                                sErrDesc = sErr
                                SalesEmployeeSync = RTN_ERROR
                                Exit Function
                            End If
                        End If

                        If oSalesEmployee_Target.GetByKey(sSalesEmployee) Then

                            SalesEmployee(oSalesEmployee_Holding, oRset.Fields.Item("Name").Value, oRset.Fields.Item("Remarks").Value)
                            ival = oSalesEmployee_Target.Update()
                            If ival <> 0 Then
                                IsError = True
                                oTragetCompany.GetLastError(iErr, sErr)
                                Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
                                sErrDesc = sErr
                                SalesEmployeeSync = RTN_ERROR
                                Exit Function
                            End If
                        Else
                            SalesEmployee(oSalesEmployee_Holding, oRset.Fields.Item("Name").Value, oRset.Fields.Item("Remarks").Value)
                            ival = oSalesEmployee_Target.Add()
                            If ival <> 0 Then
                                IsError = True
                                oTragetCompany.GetLastError(iErr, sErr)
                                Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
                                sErrDesc = sErr
                                SalesEmployeeSync = RTN_ERROR
                                Exit Function
                            End If


                        End If

                        Department(oDepartmentT, oDepartmentServicesT, oRset.Fields.Item("Name").Value, oRset.Fields.Item("Remarks").Value)
                    Else
                        sErrDesc = "No matching records found in the holding DB " & sMasterdatacode
                        SalesEmployeeSync = RTN_ERROR
                        Call WriteToLogFile("Completed with ERROR ---" & sErrDesc, sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErrDesc, sFuncName)
                        Exit Function
                    End If
                End If

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS " & sErr, sFuncName)
                SalesEmployeeSync = RTN_SUCCESS

            Catch ex As Exception
                SalesEmployeeSync = RTN_ERROR
                sErrDesc = ex.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            End Try




        End Function

        Public Function SalesEmployeeSync(ByRef oHoldingCompany As SAPbobsCOM.Company, ByRef oTragetCompany As SAPbobsCOM.Company, ByVal sMasterdatacode As String, _
                             ByVal sOSLP As String, ByRef sErrDesc As String) As Long

            'Function   :   SalesEmployeeSync()
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
            sFuncName = "SalesEmployeeSync()"
            Dim sSalesEmployee As Integer = sMasterdatacode

            Try

                Dim oSalesEmployee_Holding As SAPbobsCOM.SalesPersons = oHoldingCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oSalesPersons)
                Dim oSalesEmployee_Target As SAPbobsCOM.SalesPersons = oTragetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oSalesPersons)

                Dim oRset As SAPbobsCOM.Recordset = oHoldingCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim oCmpSrv As SAPbobsCOM.CompanyService = oTragetCompany.GetCompanyService
                Dim oDepartmentServicesT As SAPbobsCOM.DepartmentsService = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.DepartmentsService)
                Dim oDepartmentT As SAPbobsCOM.Department = oDepartmentServicesT.GetDataInterface(SAPbobsCOM.DepartmentsServiceDataInterfaces.dsDepartment)

                Dim oCmpSrvH As SAPbobsCOM.CompanyService = oHoldingCompany.GetCompanyService
                Dim oDepartmentServicesH As SAPbobsCOM.DepartmentsService = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.DepartmentsService)
                Dim oDepartmentH As SAPbobsCOM.Department = oDepartmentServicesH.GetDataInterface(SAPbobsCOM.DepartmentsServiceDataInterfaces.dsDepartment)



                If sOSLP = "S" Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to check Sales Employee is exist ", sFuncName)
                    If oSalesEmployee_Holding.GetByKey(sSalesEmployee) Then
                        Department(oDepartmentH, oDepartmentServicesH, oSalesEmployee_Holding)
                        If oSalesEmployee_Target.GetByKey(sSalesEmployee) Then

                            SalesEmployee(oSalesEmployee_Target, oSalesEmployee_Holding)
                            ival = oSalesEmployee_Target.Update()
                            If ival <> 0 Then
                                IsError = True
                                oTragetCompany.GetLastError(iErr, sErr)
                                Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
                                sErrDesc = sErr
                                SalesEmployeeSync = RTN_ERROR
                                Exit Function
                            End If

                            Department(oDepartmentT, oDepartmentServicesT, oSalesEmployee_Holding)
                        Else
                            SalesEmployee(oSalesEmployee_Target, oSalesEmployee_Holding)
                            ival = oSalesEmployee_Target.Add()
                            If ival <> 0 Then
                                IsError = True
                                oTragetCompany.GetLastError(iErr, sErr)
                                Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
                                sErrDesc = sErr
                                SalesEmployeeSync = RTN_ERROR
                                Exit Function
                            End If

                            Department(oDepartmentT, oDepartmentServicesT, oSalesEmployee_Holding)
                        End If
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS " & sErrDesc, sFuncName)

                    Else
                        sErrDesc = "No matching records found in the holding DB " & sMasterdatacode
                        SalesEmployeeSync = RTN_ERROR
                        Call WriteToLogFile("Completed with ERROR ---" & sErrDesc, sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErrDesc, sFuncName)
                        Exit Function
                    End If

                ElseIf sOSLP = "D" Then

                    sSQL = "SELECT T0.[Name], T0.[Remarks], T0.[Father] FROM OUDP T0 WHERE T0.[Name] = '" & sMasterdatacode & "'"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(" Department SQL ", sFuncName)
                    oRset.DoQuery(sSQL)
                    If oRset.RecordCount > 0 Then

                        If oSalesEmployee_Holding.GetByKey(sSalesEmployee) Then
                            SalesEmployee(oSalesEmployee_Holding, oRset.Fields.Item("Name").Value, oRset.Fields.Item("Remarks").Value)
                            ival = oSalesEmployee_Target.Update()
                            If ival <> 0 Then
                                IsError = True
                                oTragetCompany.GetLastError(iErr, sErr)
                                Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
                                sErrDesc = sErr
                                SalesEmployeeSync = RTN_ERROR
                                Exit Function
                            End If

                        Else
                            SalesEmployee(oSalesEmployee_Holding, oRset.Fields.Item("Name").Value, oRset.Fields.Item("Remarks").Value)
                            ival = oSalesEmployee_Target.Add()
                            If ival <> 0 Then
                                IsError = True
                                oTragetCompany.GetLastError(iErr, sErr)
                                Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
                                sErrDesc = sErr
                                SalesEmployeeSync = RTN_ERROR
                                Exit Function
                            End If
                        End If

                        If oSalesEmployee_Target.GetByKey(sSalesEmployee) Then

                            SalesEmployee(oSalesEmployee_Holding, oRset.Fields.Item("Name").Value, oRset.Fields.Item("Remarks").Value)
                            ival = oSalesEmployee_Target.Update()
                            If ival <> 0 Then
                                IsError = True
                                oTragetCompany.GetLastError(iErr, sErr)
                                Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
                                sErrDesc = sErr
                                SalesEmployeeSync = RTN_ERROR
                                Exit Function
                            End If
                        Else
                            SalesEmployee(oSalesEmployee_Holding, oRset.Fields.Item("Name").Value, oRset.Fields.Item("Remarks").Value)
                            ival = oSalesEmployee_Target.Add()
                            If ival <> 0 Then
                                IsError = True
                                oTragetCompany.GetLastError(iErr, sErr)
                                Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
                                sErrDesc = sErr
                                SalesEmployeeSync = RTN_ERROR
                                Exit Function
                            End If
                        End If

                        Department(oDepartmentT, oDepartmentServicesT, oRset.Fields.Item("Name").Value, oRset.Fields.Item("Remarks").Value)
                    Else
                        sErrDesc = "No matching records found in the holding DB " & sMasterdatacode
                        SalesEmployeeSync = RTN_ERROR
                        Call WriteToLogFile("Completed with ERROR ---" & sErrDesc, sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErrDesc, sFuncName)
                        Exit Function
                    End If
                End If

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS " & sErr, sFuncName)
                SalesEmployeeSync = RTN_SUCCESS
                sErrDesc = String.Empty
            Catch ex As Exception
                SalesEmployeeSync = RTN_ERROR
                sErrDesc = ex.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            End Try

        End Function

        Public Sub Department(ByRef oDepartment As SAPbobsCOM.Department, ByRef oDepartmentServices As SAPbobsCOM.DepartmentsService, ByRef oSalesEmployee As SAPbobsCOM.SalesPersons)
            oDepartment.Name = Left(Trim(oSalesEmployee.SalesEmployeeName), 20)
            oDepartment.Description = oSalesEmployee.SalesEmployeeName
            Try
                oDepartmentServices.AddDepartment(oDepartment)
            Catch ex As Exception
                ' oDepartmentServices.UpdateDepartment(oDepartment)
            End Try
        End Sub

        Public Sub Department(ByRef oDepartment As SAPbobsCOM.Department, ByRef oDepartmentServices As SAPbobsCOM.DepartmentsService, _
                             ByVal sName As String, ByVal sRemark As String)
            oDepartment.Name = Left(Trim(sName), 20)
            oDepartment.Description = sRemark
            Try
                oDepartmentServices.AddDepartment(oDepartment)
            Catch ex As Exception
                ' oDepartmentServices.UpdateDepartment(oDepartment)
            End Try
        End Sub


        Public Sub SalesEmployee(ByRef oSalesEmployee_Target As SAPbobsCOM.SalesPersons, ByRef oSalesEmployee_Holding As SAPbobsCOM.SalesPersons)
            oSalesEmployee_Target.SalesEmployeeName = oSalesEmployee_Holding.SalesEmployeeName
            oSalesEmployee_Target.Remarks = oSalesEmployee_Holding.Remarks
            If oSalesEmployee_Holding.Active = SAPbobsCOM.BoYesNoEnum.tYES Then
                oSalesEmployee_Target.Active = SAPbobsCOM.BoYesNoEnum.tYES
            Else
                oSalesEmployee_Target.Active = SAPbobsCOM.BoYesNoEnum.tNO
            End If
        End Sub

        Public Sub SalesEmployee(ByRef oSalesEmployee_Holding As SAPbobsCOM.SalesPersons, ByVal sName As String, ByVal sRemark As String)
            oSalesEmployee_Holding.SalesEmployeeName = sName
            oSalesEmployee_Holding.Remarks = sRemark
            oSalesEmployee_Holding.Active = SAPbobsCOM.BoYesNoEnum.tYES
        End Sub


    End Module

End Namespace


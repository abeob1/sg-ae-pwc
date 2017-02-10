Module modTableCreation





    Public Function UDT_UDF_Creation(ByRef sErrDesc As String) As Long


        Dim sFuncName As String = String.Empty
        Dim oCompanyTB As SAPbobsCOM.Company = Nothing
        Dim iRetValue As Integer
        Dim iErrCode As Integer


        Try
            sFuncName = "UDT_UDF_Creation"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Initializing the Company Object", sFuncName)
            Console.WriteLine("Initializing the Company Object ", sFuncName)
            oCompanyTB = New SAPbobsCOM.Company

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Assigning the representing database name", sFuncName)
            Console.WriteLine("Assigning the representing database name ", sFuncName)
            oCompanyTB.Server = p_oCompDef.sServer

            oCompanyTB.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012

            oCompanyTB.LicenseServer = p_oCompDef.sLicenseServer
            oCompanyTB.CompanyDB = p_oCompDef.sSAPDBName
            oCompanyTB.UserName = p_oCompDef.sSAPUser
            oCompanyTB.Password = p_oCompDef.sSAPPwd

            oCompanyTB.language = SAPbobsCOM.BoSuppLangs.ln_English

            oCompanyTB.UseTrusted = False

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Connecting to the Company Database.", sFuncName)
            Console.WriteLine("Connecting to the Company Database. ", sFuncName)
            iRetValue = oCompanyTB.Connect()

            If iRetValue <> 0 Then
                oCompanyTB.GetLastError(iErrCode, sErrDesc)

                sErrDesc = String.Format("Connection to Database ({0}) {1} {2} {3}", _
                    oCompanyTB.CompanyDB, System.Environment.NewLine, _
                                vbTab, sErrDesc)

                Throw New ArgumentException(sErrDesc)

                UDT_UDF_Creation = RTN_ERROR
                Exit Function
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With SUCCESS ", sFuncName)
            Console.WriteLine("Completed With SUCCESS ", sFuncName)


            AddUserTable("AE_COMPANYDATA", "AE_Company Data", SAPbobsCOM.BoUTBTableType.bott_NoObject, oCompanyTB)
            Add_Fields(oCompanyTB, sErrDesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS  ", sFuncName)
            Console.WriteLine("Completed with SUCCESS  ", sFuncName)
            UDT_UDF_Creation = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Console.WriteLine("Completed with ERROR ", sFuncName)
            UDT_UDF_Creation = RTN_ERROR
        End Try
    End Function



    Private Function AddUserTable(ByVal Name As String, ByVal Description As String, _
       ByVal Type As SAPbobsCOM.BoUTBTableType, ByRef oCompanyTB As SAPbobsCOM.Company) As Long


        Dim iRetValue As Integer
        Dim iErrCode As Integer
        Dim sErrDesc As String
        Dim sFuncName As String = "AddUserTable()"
        Try

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating a 'Company Data' Table ", sFuncName)
            Console.WriteLine("Starting Function", sFuncName)

            Dim oUserTablesMD As SAPbobsCOM.UserTablesMD

            oUserTablesMD = oCompanyTB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)

            oUserTablesMD.TableName = Name
            oUserTablesMD.TableDescription = Description
            oUserTablesMD.TableType = Type

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Create the Table  ", sFuncName)
            Console.WriteLine("Attempting to Create the Table ", sFuncName)
            iRetValue = oUserTablesMD.Add
            '// check for errors in the process
            If iRetValue <> 0 Then
                If iRetValue = -1 Then
                Else
                    oCompanyTB.GetLastError(iRetValue, sErrDesc)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR  " & sErrDesc, sFuncName)
                    Console.WriteLine("Completed with ERROR  " & sErrDesc, sFuncName)
                    If oCompanyTB.InTransaction Then
                        oCompanyTB.Disconnect()
                    End If
                    AddUserTable = RTN_ERROR
                    Exit Function
                End If
            Else

            End If
            oUserTablesMD = Nothing
            GC.Collect() 'Release the handle to the table
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS " & Name, sFuncName)
            Console.WriteLine("Completed with SUCCESS  " & Name, sFuncName)
            AddUserTable = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Console.WriteLine("Completed with ERROR ", sFuncName)
            AddUserTable = RTN_ERROR
        End Try

    End Function

    Private Sub Add_Fields(ByRef oCompanyTB As SAPbobsCOM.Company, ByRef sErrDesc As String)

        Dim lRetCode As Integer
        Dim sFuncName As String = "Add_Fields()"


        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function  ", sFuncName)
        Console.WriteLine("Starting Function ", sFuncName)
       

        Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD
        oUserFieldsMD = Nothing
        GC.Collect()
        oUserFieldsMD = oCompanyTB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        '************************************
        ' Adding "Name" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD.TableName = "@AE_COMPANYDATA"
        oUserFieldsMD.Name = "AE_UName"
        oUserFieldsMD.Description = "SAP User Name"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 40

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            oCompanyTB.GetLastError(lRetCode, sErrDesc)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR  " & sErrDesc, sFuncName)
            Console.WriteLine("Completed with ERROR  " & sErrDesc, sFuncName)
            If oCompanyTB.InTransaction Then
                oCompanyTB.Disconnect()
            End If

            Exit Sub
        End If

        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Field Created Successfully  " & oUserFieldsMD.Name, sFuncName)
        Console.WriteLine("Field Created Successfully  " & oUserFieldsMD.Name, sFuncName)


        '************************************
        ' Adding "Room" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD = oCompanyTB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_COMPANYDATA"
        oUserFieldsMD.Name = "AE_UPass"
        oUserFieldsMD.Description = "SAP User Password"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 40

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            oCompanyTB.GetLastError(lRetCode, sErrDesc)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR  " & sErrDesc, sFuncName)
            Console.WriteLine("Completed with ERROR  " & sErrDesc, sFuncName)
            If oCompanyTB.InTransaction Then
                oCompanyTB.Disconnect()
            End If
            Exit Sub
        End If
        '************************************
        ' Adding "Price" field
        '************************************
        '// Setting the Field's properties
        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Field Created Successfully  " & oUserFieldsMD.Name, sFuncName)
        Console.WriteLine("Field Created Successfully  " & oUserFieldsMD.Name, sFuncName)

        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS  ", sFuncName)
        Console.WriteLine("Completed with SUCCESS ", sFuncName)



        GC.Collect() 'Release the handle to the User Fields
    End Sub

End Module
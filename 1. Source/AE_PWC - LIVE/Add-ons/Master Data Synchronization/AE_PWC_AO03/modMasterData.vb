
Imports System.Xml
Imports System.IO


Namespace AE_PWC_AO03

    Module modMasterData

        Dim sPath As String = String.Empty
        Dim sFuncName As String = String.Empty
        Dim ival As Integer
        Dim IsError As Boolean
        Dim iErr As Integer = 0
        Dim sErr As String = String.Empty
        Dim xDoc As New XmlDocument

        Dim oMatrix As SAPbouiCOM.Matrix = Nothing

        Public Function ItemMaterSync(ByRef oHoldingCompany As SAPbobsCOM.Company, ByRef oTragetCompany As SAPbobsCOM.Company, ByVal sMasterdatacode As String, _
                                       ByRef sErrDesc As String) As Long

            'Function   :   ItemMaterSync()
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

            Try
                sFuncName = "ItemMaterSync()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function " & sMasterdatacode, sFuncName)

                Dim oItemMaster As SAPbobsCOM.Items = oHoldingCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                Dim oItemMaster_Target As SAPbobsCOM.Items = oTragetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting Item Master Sync Function ", sFuncName)

                If oItemMaster.GetByKey(sMasterdatacode) Then
                    If oItemMaster_Target.GetByKey(sMasterdatacode) Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Item_Assignment()", sFuncName)
                        'oItemMaster.SaveXML("C:\item006.xml")
                        ' oItemMaster_Target = oTragetCompany.GetBusinessObjectFromXML("C:\item006.xml", 0)
                        Item_Assignment(oItemMaster, oItemMaster_Target)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add / Update the Item Master Data " & oTragetCompany.CompanyDB, sFuncName)
                        ival = oItemMaster_Target.Update()
                        If ival <> 0 Then
                            IsError = True
                            oTragetCompany.GetLastError(iErr, sErr)
                            Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
                            sErrDesc = sErr
                            ItemMaterSync = RTN_ERROR
                            Exit Function
                        End If
                    Else
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Item_Assignment()", sFuncName)
                        oItemMaster_Target.ItemCode = oItemMaster.ItemCode
                        Item_Assignment(oItemMaster, oItemMaster_Target)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add / Update the Item Master Data " & oTragetCompany.CompanyDB, sFuncName)
                        ival = oItemMaster_Target.Add()
                        If ival <> 0 Then
                            IsError = True
                            oTragetCompany.GetLastError(iErr, sErr)
                            Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
                            sErrDesc = sErr
                            ItemMaterSync = RTN_ERROR
                            Exit Function
                        End If
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS " & sErrDesc, sFuncName)
                Else

                    sErrDesc = "No matching records found in the holding DB " & sMasterdatacode
                    ItemMaterSync = RTN_ERROR
                    Call WriteToLogFile("Completed with ERROR ---" & sErrDesc, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErrDesc, sFuncName)
                    Exit Function

                    ' 
                End If
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS ", sFuncName)
                ItemMaterSync = RTN_SUCCESS
                sErrDesc = String.Empty
            Catch ex As Exception
                ItemMaterSync = RTN_ERROR
                sErrDesc = ex.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Finally
            End Try

        End Function

        ''Public Function BPMaterSync(ByRef oHoldingCompany As SAPbobsCOM.Company, ByRef oTragetCompany As SAPbobsCOM.Company, ByVal sMasterdatacode As String, _
        ''                              ByRef sErrDesc As String) As Long

        ''    'Function   :   BPMaterSync()
        ''    'Purpose    :   
        ''    'Parameters :   ByVal oForm As SAPbouiCOM.Form
        ''    '                   oForm=Form Type
        ''    '               ByRef sErrDesc As String
        ''    '                   sErrDesc=Error Description to be returned to calling function
        ''    '               
        ''    '                   =
        ''    'Return     :   0 - FAILURE
        ''    '               1 - SUCCESS
        ''    'Author     :   SRI
        ''    'Date       :   30/12/2007
        ''    'Change     :

        ''    Dim sFuncName As String = String.Empty
        ''    Dim sBPPaymentMethods As String = String.Empty
        ''    Dim sSQLString As String = String.Empty
        ''    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
        ''    Dim oRset As SAPbobsCOM.Recordset = oTragetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        ''    Dim oDlfPaymenthod As SAPbobsCOM.Recordset = oHoldingCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        ''    Dim oBP_Holding As SAPbobsCOM.BusinessPartners = oHoldingCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
        ''    ' Dim oContact_Holding As SAPbobsCOM.PaymentRunExport = oHoldingCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPaymentRunExport)
        ''    Dim oBP_Target As SAPbobsCOM.BusinessPartners = oTragetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
        ''    '   Dim oContact_Target As SAPbobsCOM.Contacts = oTragetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oContacts)

        ''    Dim sFileName As String = System.Windows.Forms.Application.StartupPath & "\ BP.xml"
        ''    oHoldingCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
        ''    oTragetCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

        ''    Try
        ''        sFuncName = "BPMaterSync()"
        ''        If oBP_Holding.GetByKey(sMasterdatacode) Then

        ''            oBP_Holding.SaveXML(sFileName)

        ''            Dim xmlDoc As XmlDocument = New XmlDocument()
        ''            xmlDoc.Load(sFileName)

        ''            With xmlDoc.SelectSingleNode("/BOM/BO/BusinessPartners/row").CreateNavigator().AppendChild()
        ''                .WriteElementString("U_AB_SYNCCODE", sMasterdatacode)
        ''                .Close()
        ''            End With

        ''            xmlDoc.Save(sFileName)

        ''            Dim doc As New Xml.XmlDocument
        ''            doc.Load(sFileName)
        ''            Dim clientNodes = doc.SelectNodes("/BOM/BO/BPPaymentMethods/row/PaymentMethodCode")
        ''            For Each elem As Xml.XmlElement In clientNodes
        ''                If elem.InnerText = String.Empty Then
        ''                    elem.ParentNode.RemoveAll()
        ''                    Exit For
        ''                End If
        ''            Next

        ''            doc.Save(sFileName)

        ''            doc.Load(sFileName)
        ''            clientNodes = doc.SelectNodes("/BOM/BO/BPPaymentMethods/row")
        ''            For Each elem As Xml.XmlElement In clientNodes
        ''                If elem.InnerText = String.Empty Then
        ''                    elem.ParentNode.RemoveChild(elem)
        ''                    Exit For
        ''                End If
        ''            Next

        ''            doc.Save(sFileName)

        ''            sSQLString = "SELECT T0.[CardCode] FROM OCRD T0 WHERE T0.[U_AB_SYNCCODE] = '" & sMasterdatacode & "'"
        ''            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting the respective BP " & sSQLString, sFuncName)
        ''            oRset.DoQuery(sSQLString)
        ''            sMasterdatacode = oRset.Fields.Item("CardCode").Value


        ''            If oBP_Target.GetByKey(sMasterdatacode) Then

        ''                oBP_Target.Browser.ReadXml(sFileName, 0)

        ''                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Update the BP Master Data " & oTragetCompany.CompanyDB, sFuncName)
        ''                ival = oBP_Target.Update()
        ''                If ival <> 0 Then
        ''                    IsError = True
        ''                    oTragetCompany.GetLastError(iErr, sErr)
        ''                    Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
        ''                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
        ''                    sErrDesc = sErr
        ''                    BPMaterSync = RTN_ERROR
        ''                    Exit Function
        ''                End If
        ''            Else
        ''                ''  If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling BP_Assignment()", sFuncName)
        ''                Dim sCardcode As String = String.Empty
        ''                Dim oXMLDoc As New XmlDocument()
        ''                Dim oNode As XmlNode
        ''                oXMLDoc.Load(sFileName)

        ''                If oBP_Holding.Series = 2 Then
        ''                    sCardcode = oBP_Holding.CardCode
        ''                    oRset.DoQuery("SELECT T0.[CardCode] FROM OCRD T0 WHERE T0.[CardCode] ='" & sCardcode & "'")
        ''                    If oRset.RecordCount > 0 Then
        ''                        Dim sSQLstringtmp As String = "SELECT T0.[CardCode] FROM OCRD T0 WHERE T0.[CardCode] like '" & Left(sCardcode, 1) & "%' order by T0.[cardcode] DESC"
        ''                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Get the latested Cardcode " & sSQLstringtmp, sFuncName)
        ''                        oRset.DoQuery(sSQLstringtmp)
        ''                        sCardcode = oRset.Fields.Item("CardCode").Value
        ''                        sCardcode = Left(sCardcode, 1) & Right(sCardcode, sCardcode.Length - 1) + 1
        ''                    Else
        ''                        sCardcode = oBP_Holding.CardCode
        ''                    End If
        ''                Else
        ''                    sCardcode = oBP_Holding.CardCode
        ''                End If

        ''                oNode = oXMLDoc.SelectSingleNode("/BOM/BO/BusinessPartners/row/CardCode")
        ''                oNode.InnerText = sCardcode
        ''                oXMLDoc.Save(sFileName)

        ''                oBP_Target = oTragetCompany.GetBusinessObjectFromXML(sFileName, 0)
        ''                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add the BP Master Data " & oTragetCompany.CompanyDB, sFuncName)
        ''                ival = oBP_Target.Add()
        ''                If ival <> 0 Then
        ''                    IsError = True
        ''                    oTragetCompany.GetLastError(iErr, sErr)
        ''                    Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
        ''                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
        ''                    sErrDesc = sErr
        ''                    BPMaterSync = RTN_ERROR
        ''                    Exit Function
        ''                End If
        ''            End If
        ''            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS " & sErrDesc, sFuncName)
        ''        Else

        ''            sErrDesc = "No matching records found in the holding DB " & sMasterdatacode
        ''            BPMaterSync = RTN_ERROR
        ''            Call WriteToLogFile("Completed with ERROR ---" & sErrDesc, sFuncName)
        ''            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErrDesc, sFuncName)
        ''            Exit Function
        ''            ' If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS " & oTragetCompany.CompanyDB, sFuncName)
        ''        End If
        ''        BPMaterSync = RTN_SUCCESS
        ''    Catch ex As Exception
        ''        BPMaterSync = RTN_ERROR
        ''        sErrDesc = ex.Message
        ''        Call WriteToLogFile(sErrDesc, sFuncName)
        ''        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        ''    Finally
        ''        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Releasing the Objects", sFuncName)
        ''        System.Runtime.InteropServices.Marshal.ReleaseComObject(oBP_Holding)
        ''        System.Runtime.InteropServices.Marshal.ReleaseComObject(oBP_Target)
        ''        oBP_Holding = Nothing
        ''        oBP_Target = Nothing

        ''    End Try

        ''End Function

        Public Function BPMaterSync(ByRef oHoldingCompany As SAPbobsCOM.Company, ByRef oTragetCompany As SAPbobsCOM.Company, ByVal sMasterdatacode As String, _
                                      ByRef sErrDesc As String) As Long

            'Function   :   BPMaterSync()
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
            Dim sBPPaymentMethods As String = String.Empty
            Dim sSQLString As String = String.Empty

            Dim oRset_Tar As SAPbobsCOM.Recordset = Nothing
            Dim oDlfPaymenthod As SAPbobsCOM.Recordset = Nothing
            Dim oRset_Hol As SAPbobsCOM.Recordset = Nothing
            Dim oBP_Holding As SAPbobsCOM.BusinessPartners = Nothing
            Dim oBP_Holding_Banks As SAPbobsCOM.Banks = Nothing
            Dim oBP_Target_Banks As SAPbobsCOM.Banks = Nothing
            Dim sDocType As String = String.Empty

            ' Dim oContact_Holding As SAPbobsCOM.PaymentRunExport = oHoldingCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPaymentRunExport)
            Dim oBP_Target As SAPbobsCOM.BusinessPartners = Nothing
            '   Dim oContact_Target As SAPbobsCOM.Contacts = oTragetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oContacts)
            Dim bDelete As Boolean = False
            Dim sFileName As String = System.Windows.Forms.Application.StartupPath & "\ BPHolding.xml"
            Dim sFileName1 As String = System.Windows.Forms.Application.StartupPath & "\ BPTarget.xml"
            Dim sFileName2 As String = System.Windows.Forms.Application.StartupPath & "\ BPTarget1.xml"
            Dim sCurrentMsg As String = String.Empty
            Dim iAbsentryH As Integer = 0
            Dim iAbsentryT As Integer = 0
            Dim sBuyername As String = String.Empty

            Try
                sFuncName = "BPMaterSync()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

                oBP_Holding = oHoldingCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
                oBP_Target = oTragetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
                oRset_Tar = oTragetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oDlfPaymenthod = oHoldingCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRset_Hol = oHoldingCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oBP_Holding_Banks = oHoldingCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBanks)
                oBP_Target_Banks = oTragetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBanks)

                If oBP_Holding.GetByKey(sMasterdatacode) Then

                    oHoldingCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
                    oBP_Holding.SaveXML(sFileName)

                    If oBP_Holding.CardType = SAPbobsCOM.BoCardTypes.cCustomer Then
                        sDocType = "C"
                    ElseIf oBP_Holding.CardType = SAPbobsCOM.BoCardTypes.cSupplier Then
                        sDocType = "S"
                    End If

                    sSQLString = " SELECT T1.[SlpCode] FROM  " & oTragetCompany.CompanyDB & " ..[OSLP]  T1 WHERE T1.[SlpName]  = ( SELECT top(1) T1.[SlpName] FROM " & oHoldingCompany.CompanyDB & " ..OCRD T0  INNER JOIN " & oHoldingCompany.CompanyDB & " ..OSLP T1 ON T0.[SlpCode] = T1.[SlpCode] WHERE T0.[CardCode]  = '" & sMasterdatacode & "')"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting the Buyer Code " & sSQLString, sFuncName)
                    oRset_Tar.DoQuery(sSQLString)
                    sBuyername = oRset_Tar.Fields.Item("SlpCode").Value

                    sSQLString = "SELECT T0.[CardCode]  FROM OCRD T0 WHERE T0.[U_AB_SYNCCODE] = '" & sMasterdatacode & "'"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting the respective BP " & sSQLString, sFuncName)
                    oRset_Tar.DoQuery(sSQLString)
                    sMasterdatacode = oRset_Tar.Fields.Item("CardCode").Value

                    ''oHoldingCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
                    ''oBP_Holding.SaveXML(sFileName)

                    ''  /*   House Banking creation in Target Entity
                    sSQLString = "SELECT T0.[AbsEntry] FROM ODSC T0 WHERE T0.[BankCode] = '" & oBP_Holding.BPBankAccounts.BankCode & "'"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting the Holding Bank code  " & sSQLString, sFuncName)
                    oRset_Hol.DoQuery(sSQLString)
                    iAbsentryH = oRset_Hol.Fields.Item("AbsEntry").Value

                    sSQLString = "SELECT isnull(T0.[AbsEntry],0) [AbsEntry] FROM ODSC T0 WHERE T0.[BankCode] = '" & oBP_Holding.BPBankAccounts.BankCode & "'"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting the Traget Bank code " & sSQLString, sFuncName)
                    oRset_Tar.DoQuery(sSQLString)
                    iAbsentryT = oRset_Tar.Fields.Item("AbsEntry").Value

                    If oBP_Holding_Banks.GetByKey(iAbsentryH) Then
                        If Not oBP_Target_Banks.GetByKey(iAbsentryT) Then
                            oBP_Target_Banks.CountryCode = oBP_Holding_Banks.CountryCode
                            oBP_Target_Banks.BankCode = oBP_Holding_Banks.BankCode
                            oBP_Target_Banks.BankName = oBP_Holding_Banks.BankName
                            oBP_Target_Banks.SwiftNo = oBP_Holding_Banks.SwiftNo

                            Dim va As Integer = oBP_Target_Banks.Add()
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("House Bank add() value " & va, sFuncName)
                        End If

                    End If


                    If oBP_Target.GetByKey(sMasterdatacode) Then

                        ''   oBP_Target = oTragetCompany.GetBusinessObjectFromXML(sFileName, 0)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling BP_Assignment() " & sMasterdatacode, sFuncName)

                        BP_AssignmentUpdate(oBP_Holding, oBP_Target, sMasterdatacode, sCurrentMsg, True)
                        If Not String.IsNullOrEmpty(sCurrentMsg) Then
                            sErrDesc = sCurrentMsg
                        Else
                            sErrDesc = String.Empty
                        End If

                        ''======= Removing the Paymentmethods in Targeting DB
                        For imjs As Integer = oBP_Target.BPPaymentMethods.Count - 1 To 0 Step -1
                            oBP_Target.BPPaymentMethods.SetCurrentLine(imjs)
                            If Not String.IsNullOrEmpty(oBP_Target.BPPaymentMethods.PaymentMethodCode) Then
                                oBP_Target.BPPaymentMethods.Delete()
                            End If
                            If oBP_Target.BPPaymentMethods.Count = 0 Then
                                Exit For
                            End If
                        Next imjs

                        sSQLString = "SELECT T0.[PayMethCod], T0.[Descript] FROM OPYM T0 where T0.[PayMethCod] <> 'Giro'"
                        oRset_Tar.DoQuery(sSQLString)
                        For imjs As Integer = 1 To oRset_Tar.RecordCount
                            oBP_Target.BPPaymentMethods.PaymentMethodCode = oRset_Tar.Fields.Item("PayMethCod").Value
                            oBP_Target.BPPaymentMethods.Add()
                            oRset_Tar.MoveNext()
                        Next imjs

                        '' House Bank
                        sSQLString = "SELECT T0.[BankCountr], T0.[DflBnkCode], T0.[DflBnkAcct] FROM OADM T0 "
                        oRset_Tar.DoQuery(sSQLString)
                        oBP_Target.HouseBankCountry = oRset_Tar.Fields.Item("BankCountr").Value
                        oBP_Target.HouseBank = oRset_Tar.Fields.Item("DflBnkCode").Value
                        oBP_Target.HouseBankAccount = oRset_Tar.Fields.Item("DflBnkAcct").Value

                        '' Buyer name 

                        oBP_Target.SalesPersonCode = sBuyername
                        '' '' ''======= Removing the Paymentmethods in Targeting DB
                        '' ''For imjs As Integer = oBP_Target.BPPaymentMethods.Count - 1 To 0 Step -1
                        '' ''    oBP_Target.BPPaymentMethods.SetCurrentLine(imjs)
                        '' ''    If Not String.IsNullOrEmpty(oBP_Target.BPPaymentMethods.PaymentMethodCode) Then
                        '' ''        oBP_Target.BPPaymentMethods.Delete()
                        '' ''    End If
                        '' ''    If oBP_Target.BPPaymentMethods.Count = 0 Then
                        '' ''        Exit For
                        '' ''    End If
                        '' ''Next imjs

                        '' '' ''======= Adding the Paymentmethods in Targeting DB similar to Holding DB
                        '' ''For imjs As Integer = 0 To oBP_Holding.BPPaymentMethods.Count - 1
                        '' ''    oBP_Holding.BPPaymentMethods.SetCurrentLine(imjs)
                        '' ''    If Not String.IsNullOrEmpty(oBP_Holding.BPPaymentMethods.PaymentMethodCode) Then
                        '' ''        oBP_Target.BPPaymentMethods.PaymentMethodCode = oBP_Holding.BPPaymentMethods.PaymentMethodCode
                        '' ''        oBP_Target.BPPaymentMethods.Add()
                        '' ''    End If
                        '' ''Next

                        oTragetCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
                        oBP_Target.SaveXML(sFileName1)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Update the BP Master Data " & oTragetCompany.CompanyDB, sFuncName)
                        ival = oBP_Target.Update()
                        If ival <> 0 Then
                            IsError = True
                            oTragetCompany.GetLastError(iErr, sErr)
                            Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
                            sErrDesc = sErr
                            BPMaterSync = RTN_ERROR
                            Exit Function
                        Else
                            For imjs As Integer = 0 To oBP_Holding.BPBankAccounts.Count - 1
                                oBP_Holding.BPBankAccounts.SetCurrentLine(imjs)
                                If oBP_Target.DefaultBankCode = oBP_Holding.BPBankAccounts.BankCode Then
                                    If Format(oBP_Holding.BPBankAccounts.SignatureDate, "yyyyMMdd") <= "19000101" Then
                                        sSQLString = "update OCRD set [BankCtlKey] = '" & oBP_Holding.BPBankAccounts.ControlKey & "', [DflIBAN] = '" & oBP_Holding.BPBankAccounts.IBAN & "', " & _
                                       "[MandateID] = '" & oBP_Holding.BPBankAccounts.MandateID & "' FROM OCRD  WHERE CardCode = '" & oBP_Target.CardCode & "'"
                                    Else
                                        sSQLString = "update OCRD set [BankCtlKey] = '" & oBP_Holding.BPBankAccounts.ControlKey & "', [DflIBAN] = '" & oBP_Holding.BPBankAccounts.IBAN & "', " & _
                                       "[MandateID] = '" & oBP_Holding.BPBankAccounts.MandateID & "', [SignDate] = '" & Format(oBP_Holding.BPBankAccounts.SignatureDate, "yyyyMMdd") & "' FROM OCRD  WHERE CardCode = '" & oBP_Target.CardCode & "'"
                                    End If
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Updating the Sortcode " & sSQLString, sFuncName)
                                    oRset_Tar.DoQuery(sSQLString)
                                    Exit For
                                End If
                            Next
                        End If
                    Else
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling BP_Assignment()", sFuncName)
                        Dim sCardcode As String = String.Empty

                        If oBP_Holding.Series = 2 Then
                            sCardcode = oBP_Holding.CardCode
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(" Cardcode " & sCardcode, sFuncName)
                            oRset_Tar.DoQuery("SELECT T0.[CardCode] FROM OCRD T0 WHERE T0.[CardCode] ='" & sCardcode & "'")
                            If oRset_Tar.RecordCount > 0 Then
                                Dim sSQLstringtmp As String = "SELECT T0.[CardCode] FROM OCRD T0 WHERE T0.[CardCode] like '" & Left(sCardcode, 1) & "%' order by T0.[cardcode] DESC"
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Get the latested Cardcode " & sSQLstringtmp, sFuncName)
                                oRset_Tar.DoQuery(sSQLstringtmp)
                                sCardcode = oRset_Tar.Fields.Item("CardCode").Value
                                sCardcode = Left(sCardcode, 1) & Right(sCardcode, sCardcode.Length - 1) + 1
                            Else
                                sCardcode = oBP_Holding.CardCode
                            End If
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("New Cardcode is " & sCardcode, sFuncName)
                        Else
                            sCardcode = oBP_Holding.CardCode
                        End If

                        If Not String.IsNullOrEmpty(sCardcode) Then
                            oBP_Target.CardCode = sCardcode
                        End If

                        sSQLString = "SELECT series, seriesname from " & oTragetCompany.CompanyDB & " ..nnm1 where objectcode = 2 and seriesname = (SELECT seriesname from " & oHoldingCompany.CompanyDB & " ..nnm1 where objectcode = 2 and series = '" & oBP_Holding.Series & "' and DocSubType = '" & sDocType & "') and DocSubType = '" & sDocType & "' "
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(" Series " & sSQLString, sFuncName)
                        oRset_Tar.DoQuery(sSQLString)

                        oBP_Target.Series = oRset_Tar.Fields.Item("series").Value ''oBP_Holding.Series
                        BP_Assignment(oBP_Holding, oBP_Target, "", "", False)
                        oBP_Target.UserFields.Fields.Item("U_AB_SYNCCODE").Value = oBP_Holding.CardCode

                        sSQLString = "SELECT T0.[PayMethCod], T0.[Descript] FROM OPYM T0 where T0.[PayMethCod] <> 'Giro' "
                        oRset_Tar.DoQuery(sSQLString)
                        For imjs As Integer = 1 To oRset_Tar.RecordCount
                            oBP_Target.BPPaymentMethods.PaymentMethodCode = oRset_Tar.Fields.Item("PayMethCod").Value
                            oBP_Target.BPPaymentMethods.Add()
                            oRset_Tar.MoveNext()
                        Next imjs

                        '' House Bank
                        sSQLString = "SELECT T0.[BankCountr], T0.[DflBnkCode], T0.[DflBnkAcct] FROM OADM T0 "
                        oRset_Tar.DoQuery(sSQLString)
                        oBP_Target.HouseBankCountry = oRset_Tar.Fields.Item("BankCountr").Value
                        oBP_Target.HouseBank = oRset_Tar.Fields.Item("DflBnkCode").Value
                        oBP_Target.HouseBankAccount = oRset_Tar.Fields.Item("DflBnkAcct").Value
                        '' Buyer name
                        oBP_Target.SalesPersonCode = sBuyername

                        '' '' ''======= Adding the Paymentmethods in Targeting DB similiar to Holding DB
                        '' ''For imjs As Integer = 0 To oBP_Holding.BPPaymentMethods.Count - 1 Step -1
                        '' ''    oBP_Holding.BPPaymentMethods.SetCurrentLine(imjs)
                        '' ''    If Not String.IsNullOrEmpty(oBP_Holding.BPPaymentMethods.PaymentMethodCode) Then
                        '' ''        oBP_Target.BPPaymentMethods.PaymentMethodCode = oBP_Holding.BPPaymentMethods.PaymentMethodCode
                        '' ''        oBP_Target.BPPaymentMethods.Add()
                        '' ''    End If
                        '' ''Next
                        oTragetCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
                        oBP_Target.SaveXML(sFileName1)

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add the BP Master Data " & oTragetCompany.CompanyDB, sFuncName)
                        ival = oBP_Target.Add()
                        If ival <> 0 Then
                            IsError = True
                            oTragetCompany.GetLastError(iErr, sErr)
                            Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
                            sErrDesc = sErr
                            BPMaterSync = RTN_ERROR
                            Exit Function
                        Else
                            sSQLString = "SELECT T0.[CardCode] FROM OCRD T0 WHERE T0.[U_AB_SYNCCODE]  = '" & oBP_Holding.CardCode & "'"
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Finding the latest CardCode " & sSQLString, sFuncName)
                            oRset_Tar.DoQuery(sSQLString)
                            sMasterdatacode = oRset_Tar.Fields.Item(0).Value
                            If oBP_Target.GetByKey(sMasterdatacode) Then
                                For imjs As Integer = 0 To oBP_Holding.BPBankAccounts.Count - 1
                                    oBP_Holding.BPBankAccounts.SetCurrentLine(imjs)
                                    If oBP_Target.DefaultBankCode = oBP_Holding.BPBankAccounts.BankCode Then
                                        If Format(oBP_Holding.BPBankAccounts.SignatureDate, "yyyyMMdd") <= "19000101" Then
                                            sSQLString = "update OCRD set [BankCtlKey] = '" & oBP_Holding.BPBankAccounts.ControlKey & "', [DflIBAN] = '" & oBP_Holding.BPBankAccounts.IBAN & "', " & _
                                           "[MandateID] = '" & oBP_Holding.BPBankAccounts.MandateID & "' FROM OCRD  WHERE CardCode = '" & sMasterdatacode & "'"
                                        Else
                                            sSQLString = "update OCRD set [BankCtlKey] = '" & oBP_Holding.BPBankAccounts.ControlKey & "', [DflIBAN] = '" & oBP_Holding.BPBankAccounts.IBAN & "', " & _
                                           "[MandateID] = '" & oBP_Holding.BPBankAccounts.MandateID & "', [SignDate] = '" & Format(oBP_Holding.BPBankAccounts.SignatureDate, "yyyyMMdd") & "' FROM OCRD  WHERE CardCode = '" & sMasterdatacode & "'"
                                        End If
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Updating the Sortcode " & sSQLString, sFuncName)
                                        oRset_Tar.DoQuery(sSQLString)
                                        Exit For
                                    End If
                                Next
                            End If

                        End If
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS " & sErrDesc, sFuncName)
                Else

                    sErrDesc = "No matching records found in the holding DB " & sMasterdatacode
                    BPMaterSync = RTN_ERROR
                    Call WriteToLogFile("Completed with ERROR ---" & sErrDesc, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErrDesc, sFuncName)
                    Exit Function
                    ' If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS " & oTragetCompany.CompanyDB, sFuncName)
                End If
                BPMaterSync = RTN_SUCCESS
            Catch ex As Exception
                BPMaterSync = RTN_ERROR
                sErrDesc = ex.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBP_Holding)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBP_Target)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBP_Holding_Banks)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBP_Target_Banks)
                oBP_Holding = Nothing
                oBP_Target = Nothing
                oRset_Tar = Nothing
                oRset_Hol = Nothing
                oDlfPaymenthod = Nothing

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Releasing the Objects", sFuncName)
            End Try

        End Function

        Public Function BPMaterSyncOld(ByRef oHoldingCompany As SAPbobsCOM.Company, ByRef oTragetCompany As SAPbobsCOM.Company, ByVal sMasterdatacode As String, _
                                     ByRef sErrDesc As String) As Long

            'Function   :   BPMaterSync()
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

            Try
                sFuncName = "BPMaterSync()"
                Dim sBPPaymentMethods As String = String.Empty
                Dim sSQLString As String = String.Empty

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
                Dim oRset As SAPbobsCOM.Recordset = oTragetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim oBP_Holding As SAPbobsCOM.BusinessPartners = oHoldingCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
                ' Dim oContact_Holding As SAPbobsCOM.PaymentRunExport = oHoldingCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPaymentRunExport)
                Dim oBP_Target As SAPbobsCOM.BusinessPartners = oTragetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
                '   Dim oContact_Target As SAPbobsCOM.Contacts = oTragetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oContacts)

                oRset.DoQuery("SELECT T0.[DfltVendPM] FROM OADM T0")


                Dim sFileName As String = System.Windows.Forms.Application.StartupPath & "\ BP.xml"
                If oBP_Holding.GetByKey(sMasterdatacode) Then

                    ''If File.Exists(sFileName) Then
                    ''    File.Delete(sFileName)
                    ''End If
                    If oBP_Target.GetByKey(sMasterdatacode) Then

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling BP_Assignment()", sFuncName)
                        '  oBP_Target = oTragetCompany.GetBusinessObjectFromXML(sFileName, 0)
                        '' BP_Assignment(oBP_Holding, oBP_Target, sMasterdatacode)

                        For imjs As Integer = 1 To oBP_Holding.BPPaymentMethods.Count
                            oBP_Holding.BPPaymentMethods.SetCurrentLine(imjs - 1)
                            If Not String.IsNullOrEmpty(oBP_Holding.BPPaymentMethods.PaymentMethodCode) Then
                                If imjs <= oBP_Target.BPPaymentMethods.Count Then
                                    oBP_Target.BPPaymentMethods.SetCurrentLine(imjs - 1)
                                    oBP_Target.BPPaymentMethods.PaymentMethodCode = oBP_Holding.BPPaymentMethods.PaymentMethodCode
                                Else
                                    oBP_Target.BPPaymentMethods.PaymentMethodCode = oBP_Holding.BPPaymentMethods.PaymentMethodCode
                                    oBP_Target.BPPaymentMethods.Add()
                                End If
                            End If
                        Next imjs


                        For imjs As Integer = 1 To oBP_Target.BPPaymentMethods.Count
                            oBP_Target.BPPaymentMethods.SetCurrentLine(imjs - 1)
                            If String.IsNullOrEmpty(oBP_Target.BPPaymentMethods.PaymentMethodCode) Then
                                oBP_Target.BPPaymentMethods.Delete()
                            End If
                        Next imjs

                        oBP_Target.BPPaymentMethods.PaymentMethodCode = oRset.Fields.Item("DfltVendPM").Value
                        oBP_Target.BPPaymentMethods.Add()


                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add / Update the BP Master Data " & oTragetCompany.CompanyDB, sFuncName)
                        ival = oBP_Target.Update()
                        If ival <> 0 Then
                            IsError = True
                            oTragetCompany.GetLastError(iErr, sErr)
                            Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
                            sErrDesc = sErr
                            BPMaterSyncOld = RTN_ERROR
                            Exit Function
                        End If
                    Else
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Item_Assignment()", sFuncName)
                        oBP_Target.CardCode = oBP_Holding.CardCode
                        oBP_Target.Series = oBP_Holding.Series
                        ''  BP_Assignment(oBP_Holding, oBP_Target, "")
                        ''oBP_Target.BPPaymentMethods.PaymentMethodCode = oRset.Fields.Item("DfltVendPM").Value
                        ''oBP_Target.BPPaymentMethods.Add()

                        ''For imjs As Integer = 0 To oRset.RecordCount - 1
                        ''    '******RECORDSET OUTPUT COLUMN VALUE CHANGED BY JEEVA ON 07/07/2015 11:56 ISD******
                        ''    'oBP_Target.BPPaymentMethods.PaymentMethodCode = oRset.Fields.Item("PayMethCod").Value
                        ''    oBP_Target.BPPaymentMethods.PaymentMethodCode = oRset.Fields.Item("DfltVendPM").Value
                        ''    oBP_Target.BPPaymentMethods.Add()
                        ''    oRset.MoveNext()
                        ''Next
                        '' oBP_Target.SaveXML(sFileName)
                        ' oBP_Target = oTragetCompany.GetBusinessObjectFromXML(sFileName, 0)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add / Update the Item Master Data " & oTragetCompany.CompanyDB, sFuncName)
                        ival = oBP_Target.Add()
                        If ival <> 0 Then
                            IsError = True
                            oTragetCompany.GetLastError(iErr, sErr)
                            Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
                            sErrDesc = sErr
                            BPMaterSyncOld = RTN_ERROR
                            Exit Function
                        End If
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS " & sErrDesc, sFuncName)
                Else

                    sErrDesc = "No matching records found in the holding DB " & sMasterdatacode
                    BPMaterSyncOld = RTN_ERROR
                    Call WriteToLogFile("Completed with ERROR ---" & sErrDesc, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErrDesc, sFuncName)
                    Exit Function
                    ' If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS " & oTragetCompany.CompanyDB, sFuncName)
                End If
                BPMaterSyncOld = RTN_SUCCESS
            Catch ex As Exception
                BPMaterSyncOld = RTN_ERROR
                sErrDesc = ex.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Finally



            End Try

        End Function

        Public Sub Item_Assignment(ByRef oItemMaster As SAPbobsCOM.Items, ByRef oItemMaster_Target As SAPbobsCOM.Items)

            sFuncName = "Item_Assignment()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

            oItemMaster_Target.ItemName = oItemMaster.ItemName
            oItemMaster_Target.ItemType = oItemMaster.ItemType
            oItemMaster_Target.ItemsGroupCode = oItemMaster.ItemsGroupCode
            oItemMaster_Target.InventoryItem = oItemMaster.InventoryItem
            oItemMaster_Target.SalesItem = oItemMaster.SalesItem
            oItemMaster_Target.PurchaseItem = oItemMaster.PurchaseItem
            oItemMaster_Target.InventoryUOM = oItemMaster.InventoryUOM
            oItemMaster_Target.PurchaseVATGroup = oItemMaster.PurchaseVATGroup
            oItemMaster_Target.GLMethod = oItemMaster.GLMethod
            oItemMaster_Target.WTLiable = oItemMaster.WTLiable
            oItemMaster_Target.PurchaseUnit = oItemMaster.PurchaseUnit

            '  MsgBox(oItemMaster.WhsInfo.ExpensesAccount & "  - " & oItemMaster.WhsInfo.ForeignExpensAcc)
            For imjs As Integer = 0 To oItemMaster.WhsInfo.Count - 1
                oItemMaster.WhsInfo.SetCurrentLine(imjs)
                oItemMaster_Target.WhsInfo.WarehouseCode = oItemMaster.WhsInfo.WarehouseCode
                oItemMaster_Target.WhsInfo.ExpensesAccount = oItemMaster.WhsInfo.ExpensesAccount
                oItemMaster_Target.WhsInfo.ForeignExpensAcc = oItemMaster.WhsInfo.ForeignExpensAcc
                oItemMaster_Target.WhsInfo.PurchaseCreditAcc = oItemMaster.WhsInfo.PurchaseCreditAcc
                oItemMaster_Target.WhsInfo.ForeignPurchaseCreditAcc = oItemMaster.WhsInfo.ForeignPurchaseCreditAcc
                oItemMaster_Target.WhsInfo.Add()
            Next

            oItemMaster_Target.Employee = oItemMaster.Employee
            oItemMaster_Target.Properties(1) = oItemMaster.Properties(1)
            oItemMaster_Target.Properties(2) = oItemMaster.Properties(2)
            oItemMaster_Target.Properties(3) = oItemMaster.Properties(3)
            oItemMaster_Target.Properties(4) = oItemMaster.Properties(4)
            oItemMaster_Target.Properties(5) = oItemMaster.Properties(5)
            oItemMaster_Target.Properties(6) = oItemMaster.Properties(6)
            oItemMaster_Target.Properties(7) = oItemMaster.Properties(7)
            oItemMaster_Target.Properties(8) = oItemMaster.Properties(8)
            oItemMaster_Target.Properties(9) = oItemMaster.Properties(9)
            oItemMaster_Target.Properties(10) = oItemMaster.Properties(10)
            oItemMaster_Target.Properties(11) = oItemMaster.Properties(11)
            oItemMaster_Target.Properties(12) = oItemMaster.Properties(12)

            oItemMaster_Target.User_Text = oItemMaster.User_Text
            oItemMaster_Target.Frozen = SAPbobsCOM.BoYesNoEnum.tNO
            oItemMaster_Target.Valid = SAPbobsCOM.BoYesNoEnum.tYES

            'If oItemMaster.Frozen = SAPbobsCOM.BoYesNoEnum.tYES Then
            '    oItemMaster_Target.Frozen = SAPbobsCOM.BoYesNoEnum.tYES
            '    oItemMaster_Target.Valid = SAPbobsCOM.BoYesNoEnum.tNO
            'Else
            '    oItemMaster_Target.Frozen = SAPbobsCOM.BoYesNoEnum.tNO
            '    oItemMaster_Target.Valid = SAPbobsCOM.BoYesNoEnum.tYES
            'End If
            'oItemMaster_Target.Frozen = oItemMaster.Frozen
            oItemMaster_Target.FrozenFrom = oItemMaster.FrozenFrom
            oItemMaster_Target.FrozenTo = oItemMaster.FrozenTo
            oItemMaster_Target.ValidFrom = oItemMaster.ValidFrom
            oItemMaster_Target.ValidTo = oItemMaster.ValidTo

            oItemMaster_Target.UserFields.Fields.Item("U_AB_ITEMTYPE").Value = oItemMaster.UserFields.Fields.Item("U_AB_ITEMTYPE").Value
            oItemMaster_Target.UserFields.Fields.Item("U_AB_ITEMSUBGROUP").Value = oItemMaster.UserFields.Fields.Item("U_AB_ITEMSUBGROUP").Value

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

        End Sub

        Public Sub BP_Assignment(ByRef oBPMaster As SAPbobsCOM.BusinessPartners, ByRef oBPMaster_Target As SAPbobsCOM.BusinessPartners, ByVal sSynccode As String, _
                                 ByRef sCurrency As String, ByVal bUpdate As Boolean)

            sFuncName = "BP_Assignment()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)
            Dim bfound As Boolean = False

            Try
                oBPMaster_Target.CardName = oBPMaster.CardName
                '' oBPMaster_Target.Series = oBPMaster.Series
                oBPMaster_Target.CardType = oBPMaster.CardType
                oBPMaster_Target.GroupCode = oBPMaster.GroupCode
                oBPMaster_Target.GlobalLocationNumber = oBPMaster.GlobalLocationNumber
                If (oBPMaster_Target.CurrentAccountBalance = 0) Or (oBPMaster.Currency) = "##" Then
                    oBPMaster_Target.Currency = oBPMaster.Currency
                Else
                    sCurrency = "Currency not updated - Target Currency is different "
                End If

                ''  oBPMaster_Target.SalesPersonCode
                ' ''If oBPMaster.CurrentAccountBalance <> 0 Then
                ' ''    If oBPMaster.Currency <> oBPMaster_Target.Currency Then
                ' ''        sCurrency = "Currency not updated - Target Currency is different "
                ' ''    Else
                ' ''        oBPMaster_Target.Currency = oBPMaster.Currency
                ' ''    End If
                ' ''ElseIf oBPMaster_Target.CurrentAccountBalance <> 0 Then
                ' ''    If oBPMaster.Currency <> oBPMaster_Target.Currency Then
                ' ''        sCurrency = "Currency not updated - Target Currency is different "
                ' ''    Else
                ' ''        oBPMaster_Target.Currency = oBPMaster.Currency
                ' ''    End If
                ' ''Else
                ' ''    oBPMaster_Target.Currency = oBPMaster.Currency
                ' ''End If



                oBPMaster_Target.FederalTaxID = oBPMaster.FederalTaxID
                'GENERAL TAB
                oBPMaster_Target.Phone1 = oBPMaster.Phone1
                oBPMaster_Target.Phone2 = oBPMaster.Phone2
                oBPMaster_Target.Fax = oBPMaster.Fax
                oBPMaster_Target.EmailAddress = oBPMaster.EmailAddress
                '*****************ADDED ON 07/09/2015 STARTS**************
                oBPMaster_Target.Cellular = oBPMaster.Cellular
                oBPMaster_Target.Website = oBPMaster.Website
                '*****************ADDED ON 07/09/2015 ENDS**************

                'CONTACT PERSON  TAB

                ' ''For imjs As Integer = oBPMaster_Target.ContactEmployees.Count - 1 To 0 Step -1
                ' ''    oBPMaster_Target.ContactEmployees.SetCurrentLine(imjs)
                ' ''    oBPMaster_Target.ContactEmployees.Delete()
                ' ''    If oBPMaster_Target.ContactEmployees.Count = 0 Then
                ' ''        Exit For
                ' ''    End If
                ' ''Next
                If bUpdate = False Then
                    For imjs As Integer = 0 To oBPMaster.ContactEmployees.Count - 1
                        oBPMaster.ContactEmployees.SetCurrentLine(imjs)
                        ''oBPMaster_Target.ContactEmployees.Add()
                        oBPMaster_Target.ContactEmployees.Name = oBPMaster.ContactEmployees.Name
                        oBPMaster_Target.ContactEmployees.Position = oBPMaster.ContactEmployees.Position
                        oBPMaster_Target.ContactEmployees.E_Mail = oBPMaster.ContactEmployees.E_Mail
                        oBPMaster_Target.ContactEmployees.Phone1 = oBPMaster.ContactEmployees.Phone1
                        oBPMaster_Target.ContactEmployees.Phone2 = oBPMaster.ContactEmployees.Phone2

                        '*****************ADDED ON 07/09/2015 STARTS**************
                        oBPMaster_Target.ContactEmployees.FirstName = oBPMaster.ContactEmployees.FirstName
                        oBPMaster_Target.ContactEmployees.MiddleName = oBPMaster.ContactEmployees.MiddleName
                        oBPMaster_Target.ContactEmployees.LastName = oBPMaster.ContactEmployees.LastName
                        oBPMaster_Target.ContactEmployees.Title = oBPMaster.ContactEmployees.Title
                        oBPMaster_Target.ContactEmployees.Address = oBPMaster.ContactEmployees.Address
                        oBPMaster_Target.ContactEmployees.MobilePhone = oBPMaster.ContactEmployees.MobilePhone
                        oBPMaster_Target.ContactEmployees.Fax = oBPMaster.ContactEmployees.Fax
                        oBPMaster_Target.ContactEmployees.Pager = oBPMaster.ContactEmployees.Pager
                        oBPMaster_Target.ContactEmployees.Remarks1 = oBPMaster.ContactEmployees.Remarks1
                        oBPMaster_Target.ContactEmployees.Remarks2 = oBPMaster.ContactEmployees.Remarks2
                        oBPMaster_Target.ContactEmployees.Password = oBPMaster.ContactEmployees.Password
                        oBPMaster_Target.ContactEmployees.PlaceOfBirth = oBPMaster.ContactEmployees.PlaceOfBirth
                        oBPMaster_Target.ContactEmployees.DateOfBirth = oBPMaster.ContactEmployees.DateOfBirth
                        oBPMaster_Target.ContactEmployees.Gender = oBPMaster.ContactEmployees.Gender
                        oBPMaster_Target.ContactEmployees.Profession = oBPMaster.ContactEmployees.Profession
                        oBPMaster_Target.ContactEmployees.CityOfBirth = oBPMaster.ContactEmployees.CityOfBirth
                        '*****************ADDED ON 07/09/2015 ENDS**************
                        oBPMaster_Target.ContactEmployees.Add()
                    Next

                Else
                    If oBPMaster_Target.ContactEmployees.Count = 0 Then
                        For imjs As Integer = 0 To oBPMaster.ContactEmployees.Count - 1
                            oBPMaster.ContactEmployees.SetCurrentLine(imjs)
                            ''oBPMaster_Target.ContactEmployees.Add()
                            oBPMaster_Target.ContactEmployees.Name = oBPMaster.ContactEmployees.Name
                            oBPMaster_Target.ContactEmployees.Position = oBPMaster.ContactEmployees.Position
                            oBPMaster_Target.ContactEmployees.E_Mail = oBPMaster.ContactEmployees.E_Mail
                            oBPMaster_Target.ContactEmployees.Phone1 = oBPMaster.ContactEmployees.Phone1
                            oBPMaster_Target.ContactEmployees.Phone2 = oBPMaster.ContactEmployees.Phone2

                            '*****************ADDED ON 07/09/2015 STARTS**************
                            oBPMaster_Target.ContactEmployees.FirstName = oBPMaster.ContactEmployees.FirstName
                            oBPMaster_Target.ContactEmployees.MiddleName = oBPMaster.ContactEmployees.MiddleName
                            oBPMaster_Target.ContactEmployees.LastName = oBPMaster.ContactEmployees.LastName
                            oBPMaster_Target.ContactEmployees.Title = oBPMaster.ContactEmployees.Title
                            oBPMaster_Target.ContactEmployees.Address = oBPMaster.ContactEmployees.Address
                            oBPMaster_Target.ContactEmployees.MobilePhone = oBPMaster.ContactEmployees.MobilePhone
                            oBPMaster_Target.ContactEmployees.Fax = oBPMaster.ContactEmployees.Fax
                            oBPMaster_Target.ContactEmployees.Pager = oBPMaster.ContactEmployees.Pager
                            oBPMaster_Target.ContactEmployees.Remarks1 = oBPMaster.ContactEmployees.Remarks1
                            oBPMaster_Target.ContactEmployees.Remarks2 = oBPMaster.ContactEmployees.Remarks2
                            oBPMaster_Target.ContactEmployees.Password = oBPMaster.ContactEmployees.Password
                            oBPMaster_Target.ContactEmployees.PlaceOfBirth = oBPMaster.ContactEmployees.PlaceOfBirth
                            oBPMaster_Target.ContactEmployees.DateOfBirth = oBPMaster.ContactEmployees.DateOfBirth
                            oBPMaster_Target.ContactEmployees.Gender = oBPMaster.ContactEmployees.Gender
                            oBPMaster_Target.ContactEmployees.Profession = oBPMaster.ContactEmployees.Profession
                            oBPMaster_Target.ContactEmployees.CityOfBirth = oBPMaster.ContactEmployees.CityOfBirth
                            '*****************ADDED ON 07/09/2015 ENDS**************
                            oBPMaster_Target.ContactEmployees.Add()
                        Next
                    ElseIf oBPMaster_Target.ContactEmployees.Count > 0 Then
                        For imjs As Integer = 0 To oBPMaster.ContactEmployees.Count - 1
                            oBPMaster.ContactEmployees.SetCurrentLine(imjs)
                            For imjd As Integer = 0 To oBPMaster_Target.ContactEmployees.Count - 1
                                oBPMaster_Target.ContactEmployees.SetCurrentLine(imjd)
                                If oBPMaster_Target.ContactEmployees.Name = oBPMaster.ContactEmployees.Name Then
                                    oBPMaster_Target.ContactEmployees.Name = oBPMaster.ContactEmployees.Name
                                    oBPMaster_Target.ContactEmployees.Position = oBPMaster.ContactEmployees.Position
                                    oBPMaster_Target.ContactEmployees.E_Mail = oBPMaster.ContactEmployees.E_Mail
                                    oBPMaster_Target.ContactEmployees.Phone1 = oBPMaster.ContactEmployees.Phone1
                                    oBPMaster_Target.ContactEmployees.Phone2 = oBPMaster.ContactEmployees.Phone2

                                    '*****************ADDED ON 07/09/2015 STARTS**************
                                    oBPMaster_Target.ContactEmployees.FirstName = oBPMaster.ContactEmployees.FirstName
                                    oBPMaster_Target.ContactEmployees.MiddleName = oBPMaster.ContactEmployees.MiddleName
                                    oBPMaster_Target.ContactEmployees.LastName = oBPMaster.ContactEmployees.LastName
                                    oBPMaster_Target.ContactEmployees.Title = oBPMaster.ContactEmployees.Title
                                    oBPMaster_Target.ContactEmployees.Address = oBPMaster.ContactEmployees.Address
                                    oBPMaster_Target.ContactEmployees.MobilePhone = oBPMaster.ContactEmployees.MobilePhone
                                    oBPMaster_Target.ContactEmployees.Fax = oBPMaster.ContactEmployees.Fax
                                    oBPMaster_Target.ContactEmployees.Pager = oBPMaster.ContactEmployees.Pager
                                    oBPMaster_Target.ContactEmployees.Remarks1 = oBPMaster.ContactEmployees.Remarks1
                                    oBPMaster_Target.ContactEmployees.Remarks2 = oBPMaster.ContactEmployees.Remarks2
                                    oBPMaster_Target.ContactEmployees.Password = oBPMaster.ContactEmployees.Password
                                    oBPMaster_Target.ContactEmployees.PlaceOfBirth = oBPMaster.ContactEmployees.PlaceOfBirth
                                    oBPMaster_Target.ContactEmployees.DateOfBirth = oBPMaster.ContactEmployees.DateOfBirth
                                    oBPMaster_Target.ContactEmployees.Gender = oBPMaster.ContactEmployees.Gender
                                    oBPMaster_Target.ContactEmployees.Profession = oBPMaster.ContactEmployees.Profession
                                    oBPMaster_Target.ContactEmployees.CityOfBirth = oBPMaster.ContactEmployees.CityOfBirth
                                    '*****************ADDED ON 07/09/2015 ENDS**************
                                    bfound = True
                                    Exit For
                                End If
                            Next

                            If bfound = False Then
                                oBPMaster_Target.ContactEmployees.Add()
                                oBPMaster_Target.ContactEmployees.Name = oBPMaster.ContactEmployees.Name
                                oBPMaster_Target.ContactEmployees.Position = oBPMaster.ContactEmployees.Position
                                oBPMaster_Target.ContactEmployees.E_Mail = oBPMaster.ContactEmployees.E_Mail
                                oBPMaster_Target.ContactEmployees.Phone1 = oBPMaster.ContactEmployees.Phone1
                                oBPMaster_Target.ContactEmployees.Phone2 = oBPMaster.ContactEmployees.Phone2

                                '*****************ADDED ON 07/09/2015 STARTS**************
                                oBPMaster_Target.ContactEmployees.FirstName = oBPMaster.ContactEmployees.FirstName
                                oBPMaster_Target.ContactEmployees.MiddleName = oBPMaster.ContactEmployees.MiddleName
                                oBPMaster_Target.ContactEmployees.LastName = oBPMaster.ContactEmployees.LastName
                                oBPMaster_Target.ContactEmployees.Title = oBPMaster.ContactEmployees.Title
                                oBPMaster_Target.ContactEmployees.Address = oBPMaster.ContactEmployees.Address
                                oBPMaster_Target.ContactEmployees.MobilePhone = oBPMaster.ContactEmployees.MobilePhone
                                oBPMaster_Target.ContactEmployees.Fax = oBPMaster.ContactEmployees.Fax
                                oBPMaster_Target.ContactEmployees.Pager = oBPMaster.ContactEmployees.Pager
                                oBPMaster_Target.ContactEmployees.Remarks1 = oBPMaster.ContactEmployees.Remarks1
                                oBPMaster_Target.ContactEmployees.Remarks2 = oBPMaster.ContactEmployees.Remarks2
                                oBPMaster_Target.ContactEmployees.Password = oBPMaster.ContactEmployees.Password
                                oBPMaster_Target.ContactEmployees.PlaceOfBirth = oBPMaster.ContactEmployees.PlaceOfBirth
                                oBPMaster_Target.ContactEmployees.DateOfBirth = oBPMaster.ContactEmployees.DateOfBirth
                                oBPMaster_Target.ContactEmployees.Gender = oBPMaster.ContactEmployees.Gender
                                oBPMaster_Target.ContactEmployees.Profession = oBPMaster.ContactEmployees.Profession
                                oBPMaster_Target.ContactEmployees.CityOfBirth = oBPMaster.ContactEmployees.CityOfBirth
                                '*****************ADDED ON 07/09/2015 ENDS**************
                            End If
                    bfound = False
                    ''oBPMaster_Target.ContactEmployees.Add()
                        Next
                    End If
                End If


                oBPMaster_Target.ContactPerson = oBPMaster.ContactPerson

                'ADDRESS TAB 4
                For imjs As Integer = oBPMaster_Target.Addresses.Count - 1 To 0 Step -1
                    oBPMaster_Target.Addresses.SetCurrentLine(0)
                    oBPMaster_Target.Addresses.Delete()
                    If oBPMaster_Target.Addresses.Count = 0 Then
                        Exit For
                    End If
                Next

                For imjs As Integer = 0 To oBPMaster.Addresses.Count - 1
                    oBPMaster.Addresses.SetCurrentLine(imjs)
                    If Not String.IsNullOrEmpty(oBPMaster.Addresses.AddressName) Then
                        oBPMaster_Target.Addresses.AddressName = oBPMaster.Addresses.AddressName
                    End If
                    If Not String.IsNullOrEmpty(oBPMaster.Addresses.AddressName2) Then
                        oBPMaster_Target.Addresses.AddressName2 = oBPMaster.Addresses.AddressName2
                    End If
                    If Not String.IsNullOrEmpty(oBPMaster.Addresses.AddressName3) Then
                        oBPMaster_Target.Addresses.AddressName3 = oBPMaster.Addresses.AddressName3
                    End If
                    If Not String.IsNullOrEmpty(oBPMaster.Addresses.Street) Then
                        oBPMaster_Target.Addresses.Street = oBPMaster.Addresses.Street
                    End If
                    If Not String.IsNullOrEmpty(oBPMaster.Addresses.Block) Then
                        oBPMaster_Target.Addresses.Block = oBPMaster.Addresses.Block
                    End If
                    If Not String.IsNullOrEmpty(oBPMaster.Addresses.StreetNo) Then
                        oBPMaster_Target.Addresses.StreetNo = oBPMaster.Addresses.StreetNo
                    End If


                    '*****************ADDED ON 07/09/2015 STARTS**************
                    If Not String.IsNullOrEmpty(oBPMaster.Addresses.BuildingFloorRoom) Then
                        oBPMaster_Target.Addresses.BuildingFloorRoom = oBPMaster.Addresses.BuildingFloorRoom
                    End If
                    '*****************ADDED ON 07/09/2015 ENDS**************

                    If Not String.IsNullOrEmpty(oBPMaster.Addresses.City) Then
                        oBPMaster_Target.Addresses.City = oBPMaster.Addresses.City
                    End If
                    If Not String.IsNullOrEmpty(oBPMaster.Addresses.ZipCode) Then
                        oBPMaster_Target.Addresses.ZipCode = oBPMaster.Addresses.ZipCode
                    End If
                    If Not String.IsNullOrEmpty(oBPMaster.Addresses.Country) Then
                        oBPMaster_Target.Addresses.Country = oBPMaster.Addresses.Country
                    End If
                    If Not String.IsNullOrEmpty(oBPMaster.Addresses.AddressType) Then
                        oBPMaster_Target.Addresses.AddressType = oBPMaster.Addresses.AddressType
                    End If
                    oBPMaster_Target.Addresses.Add()

                Next imjs

                oBPMaster_Target.BilltoDefault = oBPMaster.BilltoDefault
                oBPMaster_Target.ShipToDefault = oBPMaster.ShipToDefault


                'PAYMENT TERMS TAB
                oBPMaster_Target.PayTermsGrpCode = oBPMaster.PayTermsGrpCode
                ''  MsgBox(oBPMaster_Target.BPBankAccounts.Count)
                '' ''If oBPMaster_Target.BPBankAccounts.Count > 0 Then
                '' ''    For imjs As Integer = oBPMaster_Target.BPBankAccounts.Count - 1 To 0 Step -1
                '' ''        oBPMaster_Target.BPBankAccounts.SetCurrentLine(0)
                '' ''        oBPMaster_Target.BPBankAccounts.Delete()
                '' ''        If oBPMaster_Target.BPBankAccounts.Count = 0 Then
                '' ''            Exit For
                '' ''        End If
                '' ''    Next
                '' ''End If

                For imjs As Integer = 0 To oBPMaster.BPBankAccounts.Count - 1
                    oBPMaster.BPBankAccounts.SetCurrentLine(imjs)
                    If Not String.IsNullOrEmpty(oBPMaster.BPBankAccounts.BankCode) Then

                        oBPMaster_Target.BPBankAccounts.Country = oBPMaster.BPBankAccounts.Country
                        oBPMaster_Target.BPBankAccounts.AccountNo = oBPMaster.BPBankAccounts.AccountNo
                        oBPMaster_Target.BPBankAccounts.BankCode = oBPMaster.BPBankAccounts.BankCode
                        If Not String.IsNullOrEmpty(sSynccode) Then
                            oBPMaster_Target.BPBankAccounts.BPCode = sSynccode 'oBPMaster.BPBankAccounts.BPCode
                        End If
                        oBPMaster_Target.BPBankAccounts.Branch = oBPMaster.BPBankAccounts.Branch
                        oBPMaster_Target.BPBankAccounts.IBAN = oBPMaster.BPBankAccounts.IBAN
                        oBPMaster_Target.BPBankAccounts.AccountName = oBPMaster.BPBankAccounts.AccountName
                        oBPMaster_Target.BPBankAccounts.BICSwiftCode = oBPMaster.BPBankAccounts.BICSwiftCode
                        oBPMaster_Target.BPBankAccounts.Street = oBPMaster.BPBankAccounts.Street
                        oBPMaster_Target.BPBankAccounts.ISRType = oBPMaster.BPBankAccounts.ISRType
                        ''****************ADDED ON 07/09/2015 STARTS**************
                        oBPMaster_Target.BPBankAccounts.MandateID = oBPMaster.BPBankAccounts.MandateID
                        oBPMaster_Target.BPBankAccounts.SignatureDate = oBPMaster.BPBankAccounts.SignatureDate
                        '' *****************ADDED ON 07/09/2015 ENDS**************
                        ''  oBPMaster_Target.BPBankAccounts.InternalKey = oBPMaster.BPBankAccounts.InternalKey
                        oBPMaster_Target.BPBankAccounts.UserNo1 = oBPMaster.BPBankAccounts.UserNo1
                        oBPMaster_Target.BPBankAccounts.UserNo2 = oBPMaster.BPBankAccounts.UserNo2
                        oBPMaster_Target.BPBankAccounts.UserNo3 = oBPMaster.BPBankAccounts.UserNo3
                        oBPMaster_Target.BPBankAccounts.UserNo4 = oBPMaster.BPBankAccounts.UserNo4
                        oBPMaster_Target.BPBankAccounts.Add()
                    End If
                Next

                '' MsgBox(oBPMaster.DefaultBankCode & " " & oBPMaster.DefaultAccount)

                ''  oBPMaster_Target.DefaultBankCode = oBPMaster.DefaultBankCode
                ''oBPMaster_Target.DefaultAccount = oBPMaster.DefaultAccount
                ''oBPMaster_Target.DefaultBranch = oBPMaster.DefaultBranch

                ' ''HOUSE BANK


                ''*****************ADDED ON 07/09/2015 STARTS**************
                oBPMaster_Target.DME = oBPMaster.DME
                oBPMaster_Target.InstructionKey = oBPMaster.InstructionKey
                ''*****************ADDED ON 07/09/2015 ENDS**************

                ' ''ACCOUNTING TAB
                If oBPMaster_Target.AccountRecivablePayables.Count = 0 Then
                    oBPMaster_Target.AccountRecivablePayables.AccountCode = oBPMaster.AccountRecivablePayables.AccountCode
                    oBPMaster_Target.AccountRecivablePayables.Add()
                Else
                    oBPMaster_Target.AccountRecivablePayables.SetCurrentLine(0)
                    oBPMaster_Target.AccountRecivablePayables.AccountCode = oBPMaster.AccountRecivablePayables.AccountCode
                End If

                oBPMaster_Target.VatLiable = oBPMaster.VatLiable
                ' oBPMaster_Target.WithholdingTaxCertified = oBPMaster.WithholdingTaxCertified
                oBPMaster_Target.VatGroup = oBPMaster.VatGroup
                oBPMaster_Target.Frozen = SAPbobsCOM.BoYesNoEnum.tNO

                oBPMaster_Target.FreeText = oBPMaster.FreeText
                oBPMaster_Target.Frozen = oBPMaster.Frozen
                If Not String.IsNullOrEmpty(oBPMaster.UserFields.Fields.Item("U_AB_WTAXREQ").Value) Then
                    oBPMaster_Target.UserFields.Fields.Item("U_AB_WTAXREQ").Value = oBPMaster.UserFields.Fields.Item("U_AB_WTAXREQ").Value
                End If

                If Not String.IsNullOrEmpty(oBPMaster.PeymentMethodCode) Then
                    oBPMaster_Target.PeymentMethodCode = oBPMaster.PeymentMethodCode
                End If

                oBPMaster_Target.Notes = oBPMaster.Notes
                oBPMaster_Target.UnifiedFederalTaxID = oBPMaster.UnifiedFederalTaxID

                oBPMaster_Target.DownPaymentClearAct = oBPMaster.DownPaymentClearAct
                oBPMaster_Target.DownPaymentInterimAccount = oBPMaster.DownPaymentInterimAccount
                oBPMaster_Target.DebitorAccount = oBPMaster.DebitorAccount

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Catch ex As Exception
                p_oSBOApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with Error " & ex.Message, sFuncName)
            End Try



        End Sub

        Public Sub BP_AssignmentUpdate(ByRef oBPMaster As SAPbobsCOM.BusinessPartners, ByRef oBPMaster_Target As SAPbobsCOM.BusinessPartners, ByVal sSynccode As String, _
                               ByRef sCurrency As String, ByVal bUpdate As Boolean)

            sFuncName = "BP_AssignmentUpdate()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)
            Dim bfound As Boolean = False

            Try
                oBPMaster_Target.CardName = oBPMaster.CardName
                '' oBPMaster_Target.Series = oBPMaster.Series
                oBPMaster_Target.CardType = oBPMaster.CardType
                oBPMaster_Target.GroupCode = oBPMaster.GroupCode
                oBPMaster_Target.GlobalLocationNumber = oBPMaster.GlobalLocationNumber
                If (oBPMaster_Target.CurrentAccountBalance = 0) Or (oBPMaster.Currency) = "##" Then
                    oBPMaster_Target.Currency = oBPMaster.Currency
                Else
                    sCurrency = "Currency not updated - Target Currency is different "
                End If

                ''  oBPMaster_Target.SalesPersonCode
                ' ''If oBPMaster.CurrentAccountBalance <> 0 Then
                ' ''    If oBPMaster.Currency <> oBPMaster_Target.Currency Then
                ' ''        sCurrency = "Currency not updated - Target Currency is different "
                ' ''    Else
                ' ''        oBPMaster_Target.Currency = oBPMaster.Currency
                ' ''    End If
                ' ''ElseIf oBPMaster_Target.CurrentAccountBalance <> 0 Then
                ' ''    If oBPMaster.Currency <> oBPMaster_Target.Currency Then
                ' ''        sCurrency = "Currency not updated - Target Currency is different "
                ' ''    Else
                ' ''        oBPMaster_Target.Currency = oBPMaster.Currency
                ' ''    End If
                ' ''Else
                ' ''    oBPMaster_Target.Currency = oBPMaster.Currency
                ' ''End If



                oBPMaster_Target.FederalTaxID = oBPMaster.FederalTaxID
                'GENERAL TAB
                oBPMaster_Target.Phone1 = oBPMaster.Phone1
                oBPMaster_Target.Phone2 = oBPMaster.Phone2
                oBPMaster_Target.Fax = oBPMaster.Fax
                oBPMaster_Target.EmailAddress = oBPMaster.EmailAddress
                '*****************ADDED ON 07/09/2015 STARTS**************
                oBPMaster_Target.Cellular = oBPMaster.Cellular
                oBPMaster_Target.Website = oBPMaster.Website
                '*****************ADDED ON 07/09/2015 ENDS**************

                'CONTACT PERSON  TAB

                ' ''For imjs As Integer = oBPMaster_Target.ContactEmployees.Count - 1 To 0 Step -1
                ' ''    oBPMaster_Target.ContactEmployees.SetCurrentLine(imjs)
                ' ''    oBPMaster_Target.ContactEmployees.Delete()
                ' ''    If oBPMaster_Target.ContactEmployees.Count = 0 Then
                ' ''        Exit For
                ' ''    End If
                ' ''Next
                If bUpdate = False Then
                    For imjs As Integer = 0 To oBPMaster.ContactEmployees.Count - 1
                        oBPMaster.ContactEmployees.SetCurrentLine(imjs)
                        ''oBPMaster_Target.ContactEmployees.Add()
                        oBPMaster_Target.ContactEmployees.Name = oBPMaster.ContactEmployees.Name
                        oBPMaster_Target.ContactEmployees.Position = oBPMaster.ContactEmployees.Position
                        oBPMaster_Target.ContactEmployees.E_Mail = oBPMaster.ContactEmployees.E_Mail
                        oBPMaster_Target.ContactEmployees.Phone1 = oBPMaster.ContactEmployees.Phone1
                        oBPMaster_Target.ContactEmployees.Phone2 = oBPMaster.ContactEmployees.Phone2

                        '*****************ADDED ON 07/09/2015 STARTS**************
                        oBPMaster_Target.ContactEmployees.FirstName = oBPMaster.ContactEmployees.FirstName
                        oBPMaster_Target.ContactEmployees.MiddleName = oBPMaster.ContactEmployees.MiddleName
                        oBPMaster_Target.ContactEmployees.LastName = oBPMaster.ContactEmployees.LastName
                        oBPMaster_Target.ContactEmployees.Title = oBPMaster.ContactEmployees.Title
                        oBPMaster_Target.ContactEmployees.Address = oBPMaster.ContactEmployees.Address
                        oBPMaster_Target.ContactEmployees.MobilePhone = oBPMaster.ContactEmployees.MobilePhone
                        oBPMaster_Target.ContactEmployees.Fax = oBPMaster.ContactEmployees.Fax
                        oBPMaster_Target.ContactEmployees.Pager = oBPMaster.ContactEmployees.Pager
                        oBPMaster_Target.ContactEmployees.Remarks1 = oBPMaster.ContactEmployees.Remarks1
                        oBPMaster_Target.ContactEmployees.Remarks2 = oBPMaster.ContactEmployees.Remarks2
                        oBPMaster_Target.ContactEmployees.Password = oBPMaster.ContactEmployees.Password
                        oBPMaster_Target.ContactEmployees.PlaceOfBirth = oBPMaster.ContactEmployees.PlaceOfBirth
                        oBPMaster_Target.ContactEmployees.DateOfBirth = oBPMaster.ContactEmployees.DateOfBirth
                        oBPMaster_Target.ContactEmployees.Gender = oBPMaster.ContactEmployees.Gender
                        oBPMaster_Target.ContactEmployees.Profession = oBPMaster.ContactEmployees.Profession
                        oBPMaster_Target.ContactEmployees.CityOfBirth = oBPMaster.ContactEmployees.CityOfBirth
                        '*****************ADDED ON 07/09/2015 ENDS**************
                        oBPMaster_Target.ContactEmployees.Add()
                    Next

                Else
                    If oBPMaster_Target.ContactEmployees.Count = 0 Then
                        For imjs As Integer = 0 To oBPMaster.ContactEmployees.Count - 1
                            oBPMaster.ContactEmployees.SetCurrentLine(imjs)
                            ''oBPMaster_Target.ContactEmployees.Add()
                            oBPMaster_Target.ContactEmployees.Name = oBPMaster.ContactEmployees.Name
                            oBPMaster_Target.ContactEmployees.Position = oBPMaster.ContactEmployees.Position
                            oBPMaster_Target.ContactEmployees.E_Mail = oBPMaster.ContactEmployees.E_Mail
                            oBPMaster_Target.ContactEmployees.Phone1 = oBPMaster.ContactEmployees.Phone1
                            oBPMaster_Target.ContactEmployees.Phone2 = oBPMaster.ContactEmployees.Phone2

                            '*****************ADDED ON 07/09/2015 STARTS**************
                            oBPMaster_Target.ContactEmployees.FirstName = oBPMaster.ContactEmployees.FirstName
                            oBPMaster_Target.ContactEmployees.MiddleName = oBPMaster.ContactEmployees.MiddleName
                            oBPMaster_Target.ContactEmployees.LastName = oBPMaster.ContactEmployees.LastName
                            oBPMaster_Target.ContactEmployees.Title = oBPMaster.ContactEmployees.Title
                            oBPMaster_Target.ContactEmployees.Address = oBPMaster.ContactEmployees.Address
                            oBPMaster_Target.ContactEmployees.MobilePhone = oBPMaster.ContactEmployees.MobilePhone
                            oBPMaster_Target.ContactEmployees.Fax = oBPMaster.ContactEmployees.Fax
                            oBPMaster_Target.ContactEmployees.Pager = oBPMaster.ContactEmployees.Pager
                            oBPMaster_Target.ContactEmployees.Remarks1 = oBPMaster.ContactEmployees.Remarks1
                            oBPMaster_Target.ContactEmployees.Remarks2 = oBPMaster.ContactEmployees.Remarks2
                            oBPMaster_Target.ContactEmployees.Password = oBPMaster.ContactEmployees.Password
                            oBPMaster_Target.ContactEmployees.PlaceOfBirth = oBPMaster.ContactEmployees.PlaceOfBirth
                            oBPMaster_Target.ContactEmployees.DateOfBirth = oBPMaster.ContactEmployees.DateOfBirth
                            oBPMaster_Target.ContactEmployees.Gender = oBPMaster.ContactEmployees.Gender
                            oBPMaster_Target.ContactEmployees.Profession = oBPMaster.ContactEmployees.Profession
                            oBPMaster_Target.ContactEmployees.CityOfBirth = oBPMaster.ContactEmployees.CityOfBirth
                            '*****************ADDED ON 07/09/2015 ENDS**************
                            oBPMaster_Target.ContactEmployees.Add()
                        Next
                    ElseIf oBPMaster_Target.ContactEmployees.Count > 0 Then
                        For imjs As Integer = 0 To oBPMaster.ContactEmployees.Count - 1
                            oBPMaster.ContactEmployees.SetCurrentLine(imjs)
                            For imjd As Integer = 0 To oBPMaster_Target.ContactEmployees.Count - 1
                                oBPMaster_Target.ContactEmployees.SetCurrentLine(imjd)
                                If oBPMaster_Target.ContactEmployees.Name = oBPMaster.ContactEmployees.Name Then
                                    oBPMaster_Target.ContactEmployees.Name = oBPMaster.ContactEmployees.Name
                                    oBPMaster_Target.ContactEmployees.Position = oBPMaster.ContactEmployees.Position
                                    oBPMaster_Target.ContactEmployees.E_Mail = oBPMaster.ContactEmployees.E_Mail
                                    oBPMaster_Target.ContactEmployees.Phone1 = oBPMaster.ContactEmployees.Phone1
                                    oBPMaster_Target.ContactEmployees.Phone2 = oBPMaster.ContactEmployees.Phone2

                                    '*****************ADDED ON 07/09/2015 STARTS**************
                                    oBPMaster_Target.ContactEmployees.FirstName = oBPMaster.ContactEmployees.FirstName
                                    oBPMaster_Target.ContactEmployees.MiddleName = oBPMaster.ContactEmployees.MiddleName
                                    oBPMaster_Target.ContactEmployees.LastName = oBPMaster.ContactEmployees.LastName
                                    oBPMaster_Target.ContactEmployees.Title = oBPMaster.ContactEmployees.Title
                                    oBPMaster_Target.ContactEmployees.Address = oBPMaster.ContactEmployees.Address
                                    oBPMaster_Target.ContactEmployees.MobilePhone = oBPMaster.ContactEmployees.MobilePhone
                                    oBPMaster_Target.ContactEmployees.Fax = oBPMaster.ContactEmployees.Fax
                                    oBPMaster_Target.ContactEmployees.Pager = oBPMaster.ContactEmployees.Pager
                                    oBPMaster_Target.ContactEmployees.Remarks1 = oBPMaster.ContactEmployees.Remarks1
                                    oBPMaster_Target.ContactEmployees.Remarks2 = oBPMaster.ContactEmployees.Remarks2
                                    oBPMaster_Target.ContactEmployees.Password = oBPMaster.ContactEmployees.Password
                                    oBPMaster_Target.ContactEmployees.PlaceOfBirth = oBPMaster.ContactEmployees.PlaceOfBirth
                                    oBPMaster_Target.ContactEmployees.DateOfBirth = oBPMaster.ContactEmployees.DateOfBirth
                                    oBPMaster_Target.ContactEmployees.Gender = oBPMaster.ContactEmployees.Gender
                                    oBPMaster_Target.ContactEmployees.Profession = oBPMaster.ContactEmployees.Profession
                                    oBPMaster_Target.ContactEmployees.CityOfBirth = oBPMaster.ContactEmployees.CityOfBirth
                                    '*****************ADDED ON 07/09/2015 ENDS**************
                                    bfound = True
                                    Exit For
                                End If
                            Next

                            If bfound = False Then
                                oBPMaster_Target.ContactEmployees.Add()
                                oBPMaster_Target.ContactEmployees.Name = oBPMaster.ContactEmployees.Name
                                oBPMaster_Target.ContactEmployees.Position = oBPMaster.ContactEmployees.Position
                                oBPMaster_Target.ContactEmployees.E_Mail = oBPMaster.ContactEmployees.E_Mail
                                oBPMaster_Target.ContactEmployees.Phone1 = oBPMaster.ContactEmployees.Phone1
                                oBPMaster_Target.ContactEmployees.Phone2 = oBPMaster.ContactEmployees.Phone2

                                '*****************ADDED ON 07/09/2015 STARTS**************
                                oBPMaster_Target.ContactEmployees.FirstName = oBPMaster.ContactEmployees.FirstName
                                oBPMaster_Target.ContactEmployees.MiddleName = oBPMaster.ContactEmployees.MiddleName
                                oBPMaster_Target.ContactEmployees.LastName = oBPMaster.ContactEmployees.LastName
                                oBPMaster_Target.ContactEmployees.Title = oBPMaster.ContactEmployees.Title
                                oBPMaster_Target.ContactEmployees.Address = oBPMaster.ContactEmployees.Address
                                oBPMaster_Target.ContactEmployees.MobilePhone = oBPMaster.ContactEmployees.MobilePhone
                                oBPMaster_Target.ContactEmployees.Fax = oBPMaster.ContactEmployees.Fax
                                oBPMaster_Target.ContactEmployees.Pager = oBPMaster.ContactEmployees.Pager
                                oBPMaster_Target.ContactEmployees.Remarks1 = oBPMaster.ContactEmployees.Remarks1
                                oBPMaster_Target.ContactEmployees.Remarks2 = oBPMaster.ContactEmployees.Remarks2
                                oBPMaster_Target.ContactEmployees.Password = oBPMaster.ContactEmployees.Password
                                oBPMaster_Target.ContactEmployees.PlaceOfBirth = oBPMaster.ContactEmployees.PlaceOfBirth
                                oBPMaster_Target.ContactEmployees.DateOfBirth = oBPMaster.ContactEmployees.DateOfBirth
                                oBPMaster_Target.ContactEmployees.Gender = oBPMaster.ContactEmployees.Gender
                                oBPMaster_Target.ContactEmployees.Profession = oBPMaster.ContactEmployees.Profession
                                oBPMaster_Target.ContactEmployees.CityOfBirth = oBPMaster.ContactEmployees.CityOfBirth
                                '*****************ADDED ON 07/09/2015 ENDS**************
                            End If
                            bfound = False
                            ''oBPMaster_Target.ContactEmployees.Add()
                        Next
                    End If
                End If


                oBPMaster_Target.ContactPerson = oBPMaster.ContactPerson

                'ADDRESS TAB 4
                For imjs As Integer = oBPMaster_Target.Addresses.Count - 1 To 0 Step -1
                    oBPMaster_Target.Addresses.SetCurrentLine(0)
                    oBPMaster_Target.Addresses.Delete()
                    If oBPMaster_Target.Addresses.Count = 0 Then
                        Exit For
                    End If
                Next

                For imjs As Integer = 0 To oBPMaster.Addresses.Count - 1
                    oBPMaster.Addresses.SetCurrentLine(imjs)
                    If Not String.IsNullOrEmpty(oBPMaster.Addresses.AddressName) Then
                        oBPMaster_Target.Addresses.AddressName = oBPMaster.Addresses.AddressName
                    End If
                    If Not String.IsNullOrEmpty(oBPMaster.Addresses.AddressName2) Then
                        oBPMaster_Target.Addresses.AddressName2 = oBPMaster.Addresses.AddressName2
                    End If
                    If Not String.IsNullOrEmpty(oBPMaster.Addresses.AddressName3) Then
                        oBPMaster_Target.Addresses.AddressName3 = oBPMaster.Addresses.AddressName3
                    End If
                    If Not String.IsNullOrEmpty(oBPMaster.Addresses.Street) Then
                        oBPMaster_Target.Addresses.Street = oBPMaster.Addresses.Street
                    End If
                    If Not String.IsNullOrEmpty(oBPMaster.Addresses.Block) Then
                        oBPMaster_Target.Addresses.Block = oBPMaster.Addresses.Block
                    End If
                    If Not String.IsNullOrEmpty(oBPMaster.Addresses.StreetNo) Then
                        oBPMaster_Target.Addresses.StreetNo = oBPMaster.Addresses.StreetNo
                    End If


                    '*****************ADDED ON 07/09/2015 STARTS**************
                    If Not String.IsNullOrEmpty(oBPMaster.Addresses.BuildingFloorRoom) Then
                        oBPMaster_Target.Addresses.BuildingFloorRoom = oBPMaster.Addresses.BuildingFloorRoom
                    End If
                    '*****************ADDED ON 07/09/2015 ENDS**************

                    If Not String.IsNullOrEmpty(oBPMaster.Addresses.City) Then
                        oBPMaster_Target.Addresses.City = oBPMaster.Addresses.City
                    End If
                    If Not String.IsNullOrEmpty(oBPMaster.Addresses.ZipCode) Then
                        oBPMaster_Target.Addresses.ZipCode = oBPMaster.Addresses.ZipCode
                    End If
                    If Not String.IsNullOrEmpty(oBPMaster.Addresses.Country) Then
                        oBPMaster_Target.Addresses.Country = oBPMaster.Addresses.Country
                    End If
                    If Not String.IsNullOrEmpty(oBPMaster.Addresses.AddressType) Then
                        oBPMaster_Target.Addresses.AddressType = oBPMaster.Addresses.AddressType
                    End If
                    oBPMaster_Target.Addresses.Add()

                Next imjs

                oBPMaster_Target.BilltoDefault = oBPMaster.BilltoDefault
                oBPMaster_Target.ShipToDefault = oBPMaster.ShipToDefault


                'PAYMENT TERMS TAB
                oBPMaster_Target.PayTermsGrpCode = oBPMaster.PayTermsGrpCode
                ''  MsgBox(oBPMaster_Target.BPBankAccounts.Count)
                '' ''If oBPMaster_Target.BPBankAccounts.Count > 0 Then
                '' ''    For imjs As Integer = oBPMaster_Target.BPBankAccounts.Count - 1 To 0 Step -1
                '' ''        oBPMaster_Target.BPBankAccounts.SetCurrentLine(0)
                '' ''        oBPMaster_Target.BPBankAccounts.Delete()
                '' ''        If oBPMaster_Target.BPBankAccounts.Count = 0 Then
                '' ''            Exit For
                '' ''        End If
                '' ''    Next
                '' ''End If

                For imjs As Integer = 0 To oBPMaster.BPBankAccounts.Count - 1
                    oBPMaster.BPBankAccounts.SetCurrentLine(imjs)
                    If Not String.IsNullOrEmpty(oBPMaster.BPBankAccounts.BankCode) Then

                        oBPMaster_Target.BPBankAccounts.Country = oBPMaster.BPBankAccounts.Country
                        oBPMaster_Target.BPBankAccounts.AccountNo = oBPMaster.BPBankAccounts.AccountNo
                        oBPMaster_Target.BPBankAccounts.BankCode = oBPMaster.BPBankAccounts.BankCode
                        If Not String.IsNullOrEmpty(sSynccode) Then
                            oBPMaster_Target.BPBankAccounts.BPCode = sSynccode 'oBPMaster.BPBankAccounts.BPCode
                        End If
                        oBPMaster_Target.BPBankAccounts.Branch = oBPMaster.BPBankAccounts.Branch
                        oBPMaster_Target.BPBankAccounts.IBAN = oBPMaster.BPBankAccounts.IBAN
                        oBPMaster_Target.BPBankAccounts.AccountName = oBPMaster.BPBankAccounts.AccountName
                        oBPMaster_Target.BPBankAccounts.BICSwiftCode = oBPMaster.BPBankAccounts.BICSwiftCode
                        oBPMaster_Target.BPBankAccounts.Street = oBPMaster.BPBankAccounts.Street
                        oBPMaster_Target.BPBankAccounts.ISRType = oBPMaster.BPBankAccounts.ISRType
                        ''**************ADDED ON 07/09/2015 STARTS**************
                        oBPMaster_Target.BPBankAccounts.MandateID = oBPMaster.BPBankAccounts.MandateID
                        oBPMaster_Target.BPBankAccounts.SignatureDate = oBPMaster.BPBankAccounts.SignatureDate
                        '' *****************ADDED ON 07/09/2015 ENDS**************
                        ''  oBPMaster_Target.BPBankAccounts.InternalKey = oBPMaster.BPBankAccounts.InternalKey
                        oBPMaster_Target.BPBankAccounts.UserNo1 = oBPMaster.BPBankAccounts.UserNo1
                        oBPMaster_Target.BPBankAccounts.UserNo2 = oBPMaster.BPBankAccounts.UserNo2
                        oBPMaster_Target.BPBankAccounts.UserNo3 = oBPMaster.BPBankAccounts.UserNo3
                        oBPMaster_Target.BPBankAccounts.UserNo4 = oBPMaster.BPBankAccounts.UserNo4
                        oBPMaster_Target.BPBankAccounts.Add()
                    Else

                        If oBPMaster_Target.BPBankAccounts.Count > 0 Then
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Count " & oBPMaster_Target.BPBankAccounts.Count, sFuncName)
                            For imjs1 As Integer = oBPMaster_Target.BPBankAccounts.Count - 1 To 0 Step -1
                                If String.IsNullOrEmpty(oBPMaster_Target.BPBankAccounts.AccountNo) Then Exit For
                                oBPMaster_Target.BPBankAccounts.SetCurrentLine(imjs1)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Delete " & imjs1, sFuncName)
                                oBPMaster_Target.BPBankAccounts.Delete()
                                If oBPMaster_Target.BPBankAccounts.Count = 0 Then
                                    Exit For
                                End If
                            Next
                        End If

                    End If
                Next

                '' MsgBox(oBPMaster.DefaultBankCode & " " & oBPMaster.DefaultAccount)

                ''  oBPMaster_Target.DefaultBankCode = oBPMaster.DefaultBankCode
                ''oBPMaster_Target.DefaultAccount = oBPMaster.DefaultAccount
                ''oBPMaster_Target.DefaultBranch = oBPMaster.DefaultBranch

                ' ''HOUSE BANK


                ''*****************ADDED ON 07/09/2015 STARTS**************
                oBPMaster_Target.DME = oBPMaster.DME
                oBPMaster_Target.InstructionKey = oBPMaster.InstructionKey
                ''*****************ADDED ON 07/09/2015 ENDS**************

                ' ''ACCOUNTING TAB
                If oBPMaster_Target.AccountRecivablePayables.Count = 0 Then
                    oBPMaster_Target.AccountRecivablePayables.AccountCode = oBPMaster.AccountRecivablePayables.AccountCode
                    oBPMaster_Target.AccountRecivablePayables.Add()
                Else
                    oBPMaster_Target.AccountRecivablePayables.SetCurrentLine(0)
                    oBPMaster_Target.AccountRecivablePayables.AccountCode = oBPMaster.AccountRecivablePayables.AccountCode
                End If

                oBPMaster_Target.VatLiable = oBPMaster.VatLiable
                ' oBPMaster_Target.WithholdingTaxCertified = oBPMaster.WithholdingTaxCertified
                oBPMaster_Target.VatGroup = oBPMaster.VatGroup
                oBPMaster_Target.Frozen = SAPbobsCOM.BoYesNoEnum.tNO

                oBPMaster_Target.FreeText = oBPMaster.FreeText
                oBPMaster_Target.Frozen = oBPMaster.Frozen
                If Not String.IsNullOrEmpty(oBPMaster.UserFields.Fields.Item("U_AB_WTAXREQ").Value) Then
                    oBPMaster_Target.UserFields.Fields.Item("U_AB_WTAXREQ").Value = oBPMaster.UserFields.Fields.Item("U_AB_WTAXREQ").Value
                End If

                If Not String.IsNullOrEmpty(oBPMaster.PeymentMethodCode) Then
                    oBPMaster_Target.PeymentMethodCode = oBPMaster.PeymentMethodCode
                End If

                oBPMaster_Target.Notes = oBPMaster.Notes
                oBPMaster_Target.UnifiedFederalTaxID = oBPMaster.UnifiedFederalTaxID

                oBPMaster_Target.DownPaymentClearAct = oBPMaster.DownPaymentClearAct
                oBPMaster_Target.DownPaymentInterimAccount = oBPMaster.DownPaymentInterimAccount
                oBPMaster_Target.DebitorAccount = oBPMaster.DebitorAccount


                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Catch ex As Exception
                p_oSBOApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with Error " & ex.Message, sFuncName)
            End Try



        End Sub


        Public Sub BP_Assignment_Update_OLD(ByRef oBPMaster As SAPbobsCOM.BusinessPartners, ByRef oBPMaster_Target As SAPbobsCOM.BusinessPartners, ByVal sSynccode As String, ByRef sCurrency As String)

            sFuncName = "BP_Assignment()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

            Try
                oBPMaster_Target.CardName = oBPMaster.CardName
                '' oBPMaster_Target.Series = oBPMaster.Series
                oBPMaster_Target.CardType = oBPMaster.CardType
                oBPMaster_Target.GroupCode = oBPMaster.GroupCode
                oBPMaster_Target.GlobalLocationNumber = oBPMaster.GlobalLocationNumber

                If oBPMaster.CurrentAccountBalance <> 0 Then
                    If oBPMaster.Currency <> oBPMaster_Target.Currency Then
                        sCurrency = "Currency not updated - Target Currency is different "
                    Else
                        oBPMaster_Target.Currency = oBPMaster.Currency
                    End If
                ElseIf oBPMaster_Target.CurrentAccountBalance <> 0 Then
                    If oBPMaster.Currency <> oBPMaster_Target.Currency Then
                        sCurrency = "Currency not updated - Target Currency is different "
                    Else
                        oBPMaster_Target.Currency = oBPMaster.Currency
                    End If
                End If

                oBPMaster_Target.FederalTaxID = oBPMaster.FederalTaxID
                'GENERAL TAB
                oBPMaster_Target.Phone1 = oBPMaster.Phone1
                oBPMaster_Target.Phone2 = oBPMaster.Phone2
                oBPMaster_Target.Fax = oBPMaster.Fax
                oBPMaster_Target.EmailAddress = oBPMaster.EmailAddress
                oBPMaster_Target.ContactPerson = oBPMaster.ContactPerson


                '*****************ADDED ON 07/09/2015 STARTS**************
                oBPMaster_Target.Cellular = oBPMaster.Cellular
                oBPMaster_Target.Website = oBPMaster.Website
                '*****************ADDED ON 07/09/2015 ENDS**************

                'CONTACT PERSON  TAB
                If oBPMaster_Target.ContactEmployees.Count = 0 Then
                    oBPMaster_Target.ContactEmployees.Add()
                    oBPMaster_Target.ContactEmployees.Name = oBPMaster.ContactEmployees.Name
                    oBPMaster_Target.ContactEmployees.Position = oBPMaster.ContactEmployees.Position
                    oBPMaster_Target.ContactEmployees.E_Mail = oBPMaster.ContactEmployees.E_Mail
                    oBPMaster_Target.ContactEmployees.Phone1 = oBPMaster.ContactEmployees.Phone1
                    oBPMaster_Target.ContactEmployees.Phone2 = oBPMaster.ContactEmployees.Phone2

                    '*****************ADDED ON 07/09/2015 STARTS**************
                    oBPMaster_Target.ContactEmployees.FirstName = oBPMaster.ContactEmployees.FirstName
                    oBPMaster_Target.ContactEmployees.MiddleName = oBPMaster.ContactEmployees.MiddleName
                    oBPMaster_Target.ContactEmployees.LastName = oBPMaster.ContactEmployees.LastName
                    oBPMaster_Target.ContactEmployees.Title = oBPMaster.ContactEmployees.Title
                    oBPMaster_Target.ContactEmployees.Address = oBPMaster.ContactEmployees.Address
                    oBPMaster_Target.ContactEmployees.MobilePhone = oBPMaster.ContactEmployees.MobilePhone
                    oBPMaster_Target.ContactEmployees.Fax = oBPMaster.ContactEmployees.Fax
                    oBPMaster_Target.ContactEmployees.Pager = oBPMaster.ContactEmployees.Pager
                    oBPMaster_Target.ContactEmployees.Remarks1 = oBPMaster.ContactEmployees.Remarks1
                    oBPMaster_Target.ContactEmployees.Remarks2 = oBPMaster.ContactEmployees.Remarks2
                    oBPMaster_Target.ContactEmployees.Password = oBPMaster.ContactEmployees.Password
                    oBPMaster_Target.ContactEmployees.PlaceOfBirth = oBPMaster.ContactEmployees.PlaceOfBirth
                    oBPMaster_Target.ContactEmployees.DateOfBirth = oBPMaster.ContactEmployees.DateOfBirth
                    oBPMaster_Target.ContactEmployees.Gender = oBPMaster.ContactEmployees.Gender
                    oBPMaster_Target.ContactEmployees.Profession = oBPMaster.ContactEmployees.Profession
                    oBPMaster_Target.ContactEmployees.CityOfBirth = oBPMaster.ContactEmployees.CityOfBirth
                    '*****************ADDED ON 07/09/2015 ENDS**************
                Else
                    ' oBPMaster_Target.ContactEmployees.Add()
                    oBPMaster_Target.ContactEmployees.SetCurrentLine(0)
                    oBPMaster_Target.ContactEmployees.Name = oBPMaster.ContactEmployees.Name
                    oBPMaster_Target.ContactEmployees.Position = oBPMaster.ContactEmployees.Position
                    oBPMaster_Target.ContactEmployees.E_Mail = oBPMaster.ContactEmployees.E_Mail
                    oBPMaster_Target.ContactEmployees.Phone1 = oBPMaster.ContactEmployees.Phone1
                    oBPMaster_Target.ContactEmployees.Phone2 = oBPMaster.ContactEmployees.Phone2

                    '*****************ADDED ON 07/09/2015 STARTS**************
                    oBPMaster_Target.ContactEmployees.FirstName = oBPMaster.ContactEmployees.FirstName
                    oBPMaster_Target.ContactEmployees.MiddleName = oBPMaster.ContactEmployees.MiddleName
                    oBPMaster_Target.ContactEmployees.LastName = oBPMaster.ContactEmployees.LastName
                    oBPMaster_Target.ContactEmployees.Title = oBPMaster.ContactEmployees.Title
                    oBPMaster_Target.ContactEmployees.Address = oBPMaster.ContactEmployees.Address
                    oBPMaster_Target.ContactEmployees.MobilePhone = oBPMaster.ContactEmployees.MobilePhone
                    oBPMaster_Target.ContactEmployees.Fax = oBPMaster.ContactEmployees.Fax
                    oBPMaster_Target.ContactEmployees.Pager = oBPMaster.ContactEmployees.Pager
                    oBPMaster_Target.ContactEmployees.Remarks1 = oBPMaster.ContactEmployees.Remarks1
                    oBPMaster_Target.ContactEmployees.Remarks2 = oBPMaster.ContactEmployees.Remarks2
                    oBPMaster_Target.ContactEmployees.Password = oBPMaster.ContactEmployees.Password
                    oBPMaster_Target.ContactEmployees.PlaceOfBirth = oBPMaster.ContactEmployees.PlaceOfBirth
                    oBPMaster_Target.ContactEmployees.DateOfBirth = oBPMaster.ContactEmployees.DateOfBirth
                    oBPMaster_Target.ContactEmployees.Gender = oBPMaster.ContactEmployees.Gender
                    oBPMaster_Target.ContactEmployees.Profession = oBPMaster.ContactEmployees.Profession
                    oBPMaster_Target.ContactEmployees.CityOfBirth = oBPMaster.ContactEmployees.CityOfBirth
                    '*****************ADDED ON 07/09/2015 ENDS**************
                End If

                'ADDRESS TAB 4
                If oBPMaster_Target.Addresses.Count = 0 Then
                    For imjs As Integer = 0 To oBPMaster.Addresses.Count - 1
                        oBPMaster_Target.Addresses.AddressName = oBPMaster.Addresses.AddressName
                        oBPMaster_Target.Addresses.AddressName2 = oBPMaster.Addresses.AddressName2
                        oBPMaster_Target.Addresses.AddressName3 = oBPMaster.Addresses.AddressName3
                        oBPMaster_Target.Addresses.Street = oBPMaster.Addresses.Street
                        oBPMaster_Target.Addresses.Block = oBPMaster.Addresses.Block

                        '*****************ADDED ON 07/09/2015 STARTS**************
                        oBPMaster_Target.Addresses.BuildingFloorRoom = oBPMaster.Addresses.BuildingFloorRoom
                        '*****************ADDED ON 07/09/2015 ENDS**************

                        oBPMaster_Target.Addresses.City = oBPMaster.Addresses.City
                        oBPMaster_Target.Addresses.ZipCode = oBPMaster.Addresses.ZipCode
                        oBPMaster_Target.Addresses.Country = oBPMaster.Addresses.Country
                        oBPMaster_Target.Addresses.AddressType = oBPMaster.Addresses.AddressType
                        oBPMaster_Target.Addresses.Add()
                    Next imjs
                Else

                    For imjs As Integer = 0 To oBPMaster.Addresses.Count - 1
                        oBPMaster.Addresses.SetCurrentLine(imjs)
                        If imjs <= oBPMaster_Target.Addresses.Count - 1 Then

                            If Not String.IsNullOrEmpty(oBPMaster.Addresses.AddressName) Then
                                oBPMaster_Target.Addresses.SetCurrentLine(imjs)
                                oBPMaster_Target.Addresses.AddressName = oBPMaster.Addresses.AddressName
                                If Not String.IsNullOrEmpty(oBPMaster.Addresses.AddressName2) Then
                                    oBPMaster_Target.Addresses.AddressName2 = oBPMaster.Addresses.AddressName2
                                End If
                                If Not String.IsNullOrEmpty(oBPMaster.Addresses.AddressName3) Then
                                    oBPMaster_Target.Addresses.AddressName3 = oBPMaster.Addresses.AddressName3
                                End If
                                If Not String.IsNullOrEmpty(oBPMaster.Addresses.Street) Then
                                    oBPMaster_Target.Addresses.Street = oBPMaster.Addresses.Street
                                End If
                                If Not String.IsNullOrEmpty(oBPMaster.Addresses.Block) Then
                                    oBPMaster_Target.Addresses.Block = oBPMaster.Addresses.Block
                                End If

                                '*****************ADDED ON 07/09/2015 STARTS**************
                                If Not String.IsNullOrEmpty(oBPMaster.Addresses.BuildingFloorRoom) Then
                                    oBPMaster_Target.Addresses.BuildingFloorRoom = oBPMaster.Addresses.BuildingFloorRoom
                                End If
                                '*****************ADDED ON 07/09/2015 ENDS**************

                                If Not String.IsNullOrEmpty(oBPMaster.Addresses.City) Then
                                    oBPMaster_Target.Addresses.City = oBPMaster.Addresses.City
                                End If
                                If Not String.IsNullOrEmpty(oBPMaster.Addresses.ZipCode) Then
                                    oBPMaster_Target.Addresses.ZipCode = oBPMaster.Addresses.ZipCode
                                End If
                                If Not String.IsNullOrEmpty(oBPMaster.Addresses.Country) Then
                                    oBPMaster_Target.Addresses.Country = oBPMaster.Addresses.Country
                                End If
                                If Not String.IsNullOrEmpty(oBPMaster.Addresses.AddressType) Then
                                    oBPMaster_Target.Addresses.AddressType = oBPMaster.Addresses.AddressType
                                End If
                            End If
                        Else
                            oBPMaster_Target.Addresses.Add()
                            oBPMaster_Target.Addresses.AddressName = oBPMaster.Addresses.AddressName
                            oBPMaster_Target.Addresses.AddressName2 = oBPMaster.Addresses.AddressName2
                            oBPMaster_Target.Addresses.AddressName3 = oBPMaster.Addresses.AddressName3
                            oBPMaster_Target.Addresses.Street = oBPMaster.Addresses.Street
                            oBPMaster_Target.Addresses.Block = oBPMaster.Addresses.Block
                            '*****************ADDED ON 07/09/2015 STARTS**************
                            oBPMaster_Target.Addresses.BuildingFloorRoom = oBPMaster.Addresses.BuildingFloorRoom
                            '*****************ADDED ON 07/09/2015 ENDS**************
                            oBPMaster_Target.Addresses.City = oBPMaster.Addresses.City
                            oBPMaster_Target.Addresses.ZipCode = oBPMaster.Addresses.ZipCode
                            oBPMaster_Target.Addresses.Country = oBPMaster.Addresses.Country
                            oBPMaster_Target.Addresses.AddressType = oBPMaster.Addresses.AddressType
                        End If
                    Next imjs
                End If

                'PAYMENT TERMS TAB
                oBPMaster_Target.PayTermsGrpCode = oBPMaster.PayTermsGrpCode

                '' ''If oBPMaster_Target.BPBankAccounts.Count > 0 Then
                '' ''    For imjs As Integer = 0 To oBPMaster_Target.BPBankAccounts.Count - 1
                '' ''        oBPMaster_Target.BPBankAccounts.SetCurrentLine(0)
                '' ''        oBPMaster_Target.BPBankAccounts.Delete()
                '' ''    Next
                '' ''End If

                '' ''If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Payment ", sFuncName)
                If oBPMaster_Target.BPBankAccounts.Count = 0 Then

                    For imjs As Integer = 0 To oBPMaster.BPBankAccounts.Count - 1
                        oBPMaster_Target.BPBankAccounts.Add()
                        oBPMaster_Target.BPBankAccounts.Country = oBPMaster.BPBankAccounts.Country
                        oBPMaster_Target.BPBankAccounts.AccountNo = oBPMaster.BPBankAccounts.AccountNo
                        oBPMaster_Target.BPBankAccounts.BankCode = oBPMaster.BPBankAccounts.BankCode
                        If Not String.IsNullOrEmpty(sSynccode) Then
                            oBPMaster_Target.BPBankAccounts.BPCode = sSynccode 'oBPMaster.BPBankAccounts.BPCode
                        End If
                        oBPMaster_Target.BPBankAccounts.Branch = oBPMaster.BPBankAccounts.Branch
                        oBPMaster_Target.BPBankAccounts.IBAN = oBPMaster.BPBankAccounts.IBAN
                        oBPMaster_Target.BPBankAccounts.AccountName = oBPMaster.BPBankAccounts.AccountName
                        oBPMaster_Target.BPBankAccounts.BICSwiftCode = oBPMaster.BPBankAccounts.BICSwiftCode

                        '*****************ADDED ON 07/09/2015 STARTS**************
                        oBPMaster_Target.BPBankAccounts.MandateID = oBPMaster.BPBankAccounts.MandateID
                        oBPMaster_Target.BPBankAccounts.SignatureDate = oBPMaster.BPBankAccounts.SignatureDate
                        '*****************ADDED ON 07/09/2015 ENDS**************
                    Next
                Else
                    For imjs As Integer = 0 To oBPMaster.BPBankAccounts.Count - 1
                        oBPMaster.BPBankAccounts.SetCurrentLine(imjs)
                        If imjs >= oBPMaster_Target.BPBankAccounts.Count Then
                            oBPMaster_Target.BPBankAccounts.Add()
                            oBPMaster_Target.BPBankAccounts.Country = oBPMaster.BPBankAccounts.Country
                            oBPMaster_Target.BPBankAccounts.AccountNo = oBPMaster.BPBankAccounts.AccountNo
                            oBPMaster_Target.BPBankAccounts.BankCode = oBPMaster.BPBankAccounts.BankCode
                            If Not String.IsNullOrEmpty(sSynccode) Then
                                oBPMaster_Target.BPBankAccounts.BPCode = sSynccode 'oBPMaster.BPBankAccounts.BPCode
                            End If
                            oBPMaster_Target.BPBankAccounts.Branch = oBPMaster.BPBankAccounts.Branch
                            oBPMaster_Target.BPBankAccounts.IBAN = oBPMaster.BPBankAccounts.IBAN
                            oBPMaster_Target.BPBankAccounts.AccountName = oBPMaster.BPBankAccounts.AccountName
                            oBPMaster_Target.BPBankAccounts.BICSwiftCode = oBPMaster.BPBankAccounts.BICSwiftCode

                            '*****************ADDED ON 07/09/2015 STARTS**************
                            oBPMaster_Target.BPBankAccounts.MandateID = oBPMaster.BPBankAccounts.MandateID
                            oBPMaster_Target.BPBankAccounts.SignatureDate = oBPMaster.BPBankAccounts.SignatureDate
                            '*****************ADDED ON 07/09/2015 ENDS**************
                        Else
                            For imjd As Integer = 0 To oBPMaster_Target.BPBankAccounts.Count - 1
                                oBPMaster_Target.BPBankAccounts.SetCurrentLine(imjd)
                                ''   MsgBox("target " & oBPMaster_Target.BPBankAccounts.AccountNo & " holding " & oBPMaster.BPBankAccounts.AccountNo & "Internal base " & oBPMaster.BPBankAccounts.InternalKey & " target " & oBPMaster_Target.BPBankAccounts.InternalKey)
                                If Not String.IsNullOrEmpty(oBPMaster.BPBankAccounts.AccountNo) Then
                                    If oBPMaster_Target.BPBankAccounts.AccountNo = oBPMaster.BPBankAccounts.AccountNo Then
                                        oBPMaster_Target.BPBankAccounts.Country = oBPMaster.BPBankAccounts.Country
                                        oBPMaster_Target.BPBankAccounts.AccountNo = oBPMaster.BPBankAccounts.AccountNo
                                        oBPMaster_Target.BPBankAccounts.BankCode = oBPMaster.BPBankAccounts.BankCode
                                        If Not String.IsNullOrEmpty(sSynccode) Then
                                            oBPMaster_Target.BPBankAccounts.BPCode = sSynccode 'oBPMaster.BPBankAccounts.BPCode
                                        End If
                                        oBPMaster_Target.BPBankAccounts.Branch = oBPMaster.BPBankAccounts.Branch
                                        oBPMaster_Target.BPBankAccounts.IBAN = oBPMaster.BPBankAccounts.IBAN
                                        oBPMaster_Target.BPBankAccounts.AccountName = oBPMaster.BPBankAccounts.AccountName
                                        oBPMaster_Target.BPBankAccounts.BICSwiftCode = oBPMaster.BPBankAccounts.BICSwiftCode

                                        '*****************ADDED ON 07/09/2015 STARTS**************
                                        oBPMaster_Target.BPBankAccounts.MandateID = oBPMaster.BPBankAccounts.MandateID
                                        oBPMaster_Target.BPBankAccounts.SignatureDate = oBPMaster.BPBankAccounts.SignatureDate
                                    End If
                                End If

                            Next
                        End If
                    Next
                End If

                '' oBPMaster_Target.SaveXML("E:\Abeo-Projects\PWC\SVN - Copy\1. Source\AE_PWC\Add-ons\Master Data Synchronization\AE_PWC_AO03\bin\Debug\BP.xml")
                ' ''PAYMENT RUN TAB
                oBPMaster_Target.HouseBankCountry = oBPMaster.HouseBankCountry
                oBPMaster_Target.HouseBank = oBPMaster.HouseBank
                oBPMaster_Target.HouseBankAccount = oBPMaster.HouseBankAccount

                ''*****************ADDED ON 07/09/2015 STARTS**************
                oBPMaster_Target.DME = oBPMaster.DME
                oBPMaster_Target.InstructionKey = oBPMaster.InstructionKey
                ''*****************ADDED ON 07/09/2015 ENDS**************
                ' ''ACCOUNTING TAB
                If oBPMaster_Target.AccountRecivablePayables.Count = 0 Then
                    oBPMaster_Target.AccountRecivablePayables.AccountCode = oBPMaster.AccountRecivablePayables.AccountCode
                    oBPMaster_Target.AccountRecivablePayables.Add()
                Else
                    oBPMaster_Target.AccountRecivablePayables.SetCurrentLine(0)
                    oBPMaster_Target.AccountRecivablePayables.AccountCode = oBPMaster.AccountRecivablePayables.AccountCode
                End If

                oBPMaster_Target.VatLiable = oBPMaster.VatLiable
                ' oBPMaster_Target.WithholdingTaxCertified = oBPMaster.WithholdingTaxCertified
                oBPMaster_Target.VatGroup = oBPMaster.VatGroup
                oBPMaster_Target.Frozen = SAPbobsCOM.BoYesNoEnum.tNO

                oBPMaster_Target.FreeText = oBPMaster.FreeText
                oBPMaster_Target.Frozen = oBPMaster.Frozen
                If Not String.IsNullOrEmpty(oBPMaster.UserFields.Fields.Item("U_AB_WTAXREQ").Value) Then
                    oBPMaster_Target.UserFields.Fields.Item("U_AB_WTAXREQ").Value = oBPMaster.UserFields.Fields.Item("U_AB_WTAXREQ").Value
                End If

                If Not String.IsNullOrEmpty(oBPMaster.PeymentMethodCode) Then
                    oBPMaster_Target.PeymentMethodCode = oBPMaster.PeymentMethodCode
                End If
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Catch ex As Exception
                p_oSBOApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with Error " & ex.Message, sFuncName)
            End Try



        End Sub

        Public Sub BP_Assignment_03112015(ByRef oBPMaster As SAPbobsCOM.BusinessPartners, ByRef oBPMaster_Target As SAPbobsCOM.BusinessPartners, ByVal sSynccode As String, ByRef sCurrency As String)

            sFuncName = "BP_Assignment()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

            Try
                oBPMaster_Target.CardName = oBPMaster.CardName
                '' oBPMaster_Target.Series = oBPMaster.Series
                oBPMaster_Target.CardType = oBPMaster.CardType
                oBPMaster_Target.GroupCode = oBPMaster.GroupCode
                oBPMaster_Target.GlobalLocationNumber = oBPMaster.GlobalLocationNumber

                If oBPMaster.CurrentAccountBalance <> 0 Then
                    If oBPMaster.Currency <> oBPMaster_Target.Currency Then
                        sCurrency = "Currency not updated - Target Currency is different "
                    Else
                        oBPMaster_Target.Currency = oBPMaster.Currency
                    End If
                ElseIf oBPMaster_Target.CurrentAccountBalance <> 0 Then
                    If oBPMaster.Currency <> oBPMaster_Target.Currency Then
                        sCurrency = "Currency not updated - Target Currency is different "
                    Else
                        oBPMaster_Target.Currency = oBPMaster.Currency
                    End If
                End If

                oBPMaster_Target.FederalTaxID = oBPMaster.FederalTaxID
                'GENERAL TAB
                oBPMaster_Target.Phone1 = oBPMaster.Phone1
                oBPMaster_Target.Phone2 = oBPMaster.Phone2
                oBPMaster_Target.Fax = oBPMaster.Fax
                oBPMaster_Target.EmailAddress = oBPMaster.EmailAddress
                oBPMaster_Target.ContactPerson = oBPMaster.ContactPerson


                '*****************ADDED ON 07/09/2015 STARTS**************
                oBPMaster_Target.Cellular = oBPMaster.Cellular
                oBPMaster_Target.Website = oBPMaster.Website
                '*****************ADDED ON 07/09/2015 ENDS**************

                'CONTACT PERSON  TAB
                If oBPMaster_Target.ContactEmployees.Count = 0 Then
                    oBPMaster_Target.ContactEmployees.Add()
                    oBPMaster_Target.ContactEmployees.Name = oBPMaster.ContactEmployees.Name
                    oBPMaster_Target.ContactEmployees.Position = oBPMaster.ContactEmployees.Position
                    oBPMaster_Target.ContactEmployees.E_Mail = oBPMaster.ContactEmployees.E_Mail
                    oBPMaster_Target.ContactEmployees.Phone1 = oBPMaster.ContactEmployees.Phone1
                    oBPMaster_Target.ContactEmployees.Phone2 = oBPMaster.ContactEmployees.Phone2

                    '*****************ADDED ON 07/09/2015 STARTS**************
                    oBPMaster_Target.ContactEmployees.FirstName = oBPMaster.ContactEmployees.FirstName
                    oBPMaster_Target.ContactEmployees.MiddleName = oBPMaster.ContactEmployees.MiddleName
                    oBPMaster_Target.ContactEmployees.LastName = oBPMaster.ContactEmployees.LastName
                    oBPMaster_Target.ContactEmployees.Title = oBPMaster.ContactEmployees.Title
                    oBPMaster_Target.ContactEmployees.Address = oBPMaster.ContactEmployees.Address
                    oBPMaster_Target.ContactEmployees.MobilePhone = oBPMaster.ContactEmployees.MobilePhone
                    oBPMaster_Target.ContactEmployees.Fax = oBPMaster.ContactEmployees.Fax
                    oBPMaster_Target.ContactEmployees.Pager = oBPMaster.ContactEmployees.Pager
                    oBPMaster_Target.ContactEmployees.Remarks1 = oBPMaster.ContactEmployees.Remarks1
                    oBPMaster_Target.ContactEmployees.Remarks2 = oBPMaster.ContactEmployees.Remarks2
                    oBPMaster_Target.ContactEmployees.Password = oBPMaster.ContactEmployees.Password
                    oBPMaster_Target.ContactEmployees.PlaceOfBirth = oBPMaster.ContactEmployees.PlaceOfBirth
                    oBPMaster_Target.ContactEmployees.DateOfBirth = oBPMaster.ContactEmployees.DateOfBirth
                    oBPMaster_Target.ContactEmployees.Gender = oBPMaster.ContactEmployees.Gender
                    oBPMaster_Target.ContactEmployees.Profession = oBPMaster.ContactEmployees.Profession
                    oBPMaster_Target.ContactEmployees.CityOfBirth = oBPMaster.ContactEmployees.CityOfBirth
                    '*****************ADDED ON 07/09/2015 ENDS**************
                Else
                    ' oBPMaster_Target.ContactEmployees.Add()
                    oBPMaster_Target.ContactEmployees.SetCurrentLine(0)
                    oBPMaster_Target.ContactEmployees.Name = oBPMaster.ContactEmployees.Name
                    oBPMaster_Target.ContactEmployees.Position = oBPMaster.ContactEmployees.Position
                    oBPMaster_Target.ContactEmployees.E_Mail = oBPMaster.ContactEmployees.E_Mail
                    oBPMaster_Target.ContactEmployees.Phone1 = oBPMaster.ContactEmployees.Phone1
                    oBPMaster_Target.ContactEmployees.Phone2 = oBPMaster.ContactEmployees.Phone2

                    '*****************ADDED ON 07/09/2015 STARTS**************
                    oBPMaster_Target.ContactEmployees.FirstName = oBPMaster.ContactEmployees.FirstName
                    oBPMaster_Target.ContactEmployees.MiddleName = oBPMaster.ContactEmployees.MiddleName
                    oBPMaster_Target.ContactEmployees.LastName = oBPMaster.ContactEmployees.LastName
                    oBPMaster_Target.ContactEmployees.Title = oBPMaster.ContactEmployees.Title
                    oBPMaster_Target.ContactEmployees.Address = oBPMaster.ContactEmployees.Address
                    oBPMaster_Target.ContactEmployees.MobilePhone = oBPMaster.ContactEmployees.MobilePhone
                    oBPMaster_Target.ContactEmployees.Fax = oBPMaster.ContactEmployees.Fax
                    oBPMaster_Target.ContactEmployees.Pager = oBPMaster.ContactEmployees.Pager
                    oBPMaster_Target.ContactEmployees.Remarks1 = oBPMaster.ContactEmployees.Remarks1
                    oBPMaster_Target.ContactEmployees.Remarks2 = oBPMaster.ContactEmployees.Remarks2
                    oBPMaster_Target.ContactEmployees.Password = oBPMaster.ContactEmployees.Password
                    oBPMaster_Target.ContactEmployees.PlaceOfBirth = oBPMaster.ContactEmployees.PlaceOfBirth
                    oBPMaster_Target.ContactEmployees.DateOfBirth = oBPMaster.ContactEmployees.DateOfBirth
                    oBPMaster_Target.ContactEmployees.Gender = oBPMaster.ContactEmployees.Gender
                    oBPMaster_Target.ContactEmployees.Profession = oBPMaster.ContactEmployees.Profession
                    oBPMaster_Target.ContactEmployees.CityOfBirth = oBPMaster.ContactEmployees.CityOfBirth
                    '*****************ADDED ON 07/09/2015 ENDS**************
                End If

                'ADDRESS TAB 4
                If oBPMaster_Target.Addresses.Count = 0 Then
                    For imjs As Integer = 0 To oBPMaster.Addresses.Count - 1
                        oBPMaster_Target.Addresses.AddressName = oBPMaster.Addresses.AddressName
                        oBPMaster_Target.Addresses.AddressName2 = oBPMaster.Addresses.AddressName2
                        oBPMaster_Target.Addresses.AddressName3 = oBPMaster.Addresses.AddressName3
                        oBPMaster_Target.Addresses.Street = oBPMaster.Addresses.Street
                        oBPMaster_Target.Addresses.Block = oBPMaster.Addresses.Block

                        '*****************ADDED ON 07/09/2015 STARTS**************
                        oBPMaster_Target.Addresses.BuildingFloorRoom = oBPMaster.Addresses.BuildingFloorRoom
                        '*****************ADDED ON 07/09/2015 ENDS**************

                        oBPMaster_Target.Addresses.City = oBPMaster.Addresses.City
                        oBPMaster_Target.Addresses.ZipCode = oBPMaster.Addresses.ZipCode
                        oBPMaster_Target.Addresses.Country = oBPMaster.Addresses.Country
                        oBPMaster_Target.Addresses.AddressType = oBPMaster.Addresses.AddressType
                        oBPMaster_Target.Addresses.Add()
                    Next imjs
                Else

                    For imjs As Integer = 0 To oBPMaster.Addresses.Count - 1
                        oBPMaster.Addresses.SetCurrentLine(imjs)
                        If imjs <= oBPMaster_Target.Addresses.Count - 1 Then

                            If Not String.IsNullOrEmpty(oBPMaster.Addresses.AddressName) Then
                                oBPMaster_Target.Addresses.SetCurrentLine(imjs)
                                oBPMaster_Target.Addresses.AddressName = oBPMaster.Addresses.AddressName
                                If Not String.IsNullOrEmpty(oBPMaster.Addresses.AddressName2) Then
                                    oBPMaster_Target.Addresses.AddressName2 = oBPMaster.Addresses.AddressName2
                                End If
                                If Not String.IsNullOrEmpty(oBPMaster.Addresses.AddressName3) Then
                                    oBPMaster_Target.Addresses.AddressName3 = oBPMaster.Addresses.AddressName3
                                End If
                                If Not String.IsNullOrEmpty(oBPMaster.Addresses.Street) Then
                                    oBPMaster_Target.Addresses.Street = oBPMaster.Addresses.Street
                                End If
                                If Not String.IsNullOrEmpty(oBPMaster.Addresses.Block) Then
                                    oBPMaster_Target.Addresses.Block = oBPMaster.Addresses.Block
                                End If

                                '*****************ADDED ON 07/09/2015 STARTS**************
                                If Not String.IsNullOrEmpty(oBPMaster.Addresses.BuildingFloorRoom) Then
                                    oBPMaster_Target.Addresses.BuildingFloorRoom = oBPMaster.Addresses.BuildingFloorRoom
                                End If
                                '*****************ADDED ON 07/09/2015 ENDS**************

                                If Not String.IsNullOrEmpty(oBPMaster.Addresses.City) Then
                                    oBPMaster_Target.Addresses.City = oBPMaster.Addresses.City
                                End If
                                If Not String.IsNullOrEmpty(oBPMaster.Addresses.ZipCode) Then
                                    oBPMaster_Target.Addresses.ZipCode = oBPMaster.Addresses.ZipCode
                                End If
                                If Not String.IsNullOrEmpty(oBPMaster.Addresses.Country) Then
                                    oBPMaster_Target.Addresses.Country = oBPMaster.Addresses.Country
                                End If
                                If Not String.IsNullOrEmpty(oBPMaster.Addresses.AddressType) Then
                                    oBPMaster_Target.Addresses.AddressType = oBPMaster.Addresses.AddressType
                                End If
                            End If
                        Else
                            oBPMaster_Target.Addresses.Add()
                            oBPMaster_Target.Addresses.AddressName = oBPMaster.Addresses.AddressName
                            oBPMaster_Target.Addresses.AddressName2 = oBPMaster.Addresses.AddressName2
                            oBPMaster_Target.Addresses.AddressName3 = oBPMaster.Addresses.AddressName3
                            oBPMaster_Target.Addresses.Street = oBPMaster.Addresses.Street
                            oBPMaster_Target.Addresses.Block = oBPMaster.Addresses.Block
                            '*****************ADDED ON 07/09/2015 STARTS**************
                            oBPMaster_Target.Addresses.BuildingFloorRoom = oBPMaster.Addresses.BuildingFloorRoom
                            '*****************ADDED ON 07/09/2015 ENDS**************
                            oBPMaster_Target.Addresses.City = oBPMaster.Addresses.City
                            oBPMaster_Target.Addresses.ZipCode = oBPMaster.Addresses.ZipCode
                            oBPMaster_Target.Addresses.Country = oBPMaster.Addresses.Country
                            oBPMaster_Target.Addresses.AddressType = oBPMaster.Addresses.AddressType
                        End If
                    Next imjs
                End If

                'PAYMENT TERMS TAB
                oBPMaster_Target.PayTermsGrpCode = oBPMaster.PayTermsGrpCode

                '' ''If oBPMaster_Target.BPBankAccounts.Count > 0 Then
                '' ''    For imjs As Integer = 0 To oBPMaster_Target.BPBankAccounts.Count - 1
                '' ''        oBPMaster_Target.BPBankAccounts.SetCurrentLine(0)
                '' ''        oBPMaster_Target.BPBankAccounts.Delete()
                '' ''    Next
                '' ''End If

                '' ''If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Payment ", sFuncName)
                If oBPMaster_Target.BPBankAccounts.Count = 0 Then

                    For imjs As Integer = 0 To oBPMaster.BPBankAccounts.Count - 1
                        oBPMaster_Target.BPBankAccounts.Add()
                        oBPMaster_Target.BPBankAccounts.Country = oBPMaster.BPBankAccounts.Country
                        oBPMaster_Target.BPBankAccounts.AccountNo = oBPMaster.BPBankAccounts.AccountNo
                        oBPMaster_Target.BPBankAccounts.BankCode = oBPMaster.BPBankAccounts.BankCode
                        If Not String.IsNullOrEmpty(sSynccode) Then
                            oBPMaster_Target.BPBankAccounts.BPCode = sSynccode 'oBPMaster.BPBankAccounts.BPCode
                        End If
                        oBPMaster_Target.BPBankAccounts.Branch = oBPMaster.BPBankAccounts.Branch
                        oBPMaster_Target.BPBankAccounts.IBAN = oBPMaster.BPBankAccounts.IBAN
                        oBPMaster_Target.BPBankAccounts.AccountName = oBPMaster.BPBankAccounts.AccountName
                        oBPMaster_Target.BPBankAccounts.BICSwiftCode = oBPMaster.BPBankAccounts.BICSwiftCode

                        '*****************ADDED ON 07/09/2015 STARTS**************
                        oBPMaster_Target.BPBankAccounts.MandateID = oBPMaster.BPBankAccounts.MandateID
                        oBPMaster_Target.BPBankAccounts.SignatureDate = oBPMaster.BPBankAccounts.SignatureDate
                        '*****************ADDED ON 07/09/2015 ENDS**************
                    Next
                Else
                    For imjs As Integer = 0 To oBPMaster.BPBankAccounts.Count - 1
                        oBPMaster.BPBankAccounts.SetCurrentLine(imjs)
                        If imjs >= oBPMaster_Target.BPBankAccounts.Count Then
                            oBPMaster_Target.BPBankAccounts.Add()
                            oBPMaster_Target.BPBankAccounts.Country = oBPMaster.BPBankAccounts.Country
                            oBPMaster_Target.BPBankAccounts.AccountNo = oBPMaster.BPBankAccounts.AccountNo
                            oBPMaster_Target.BPBankAccounts.BankCode = oBPMaster.BPBankAccounts.BankCode
                            If Not String.IsNullOrEmpty(sSynccode) Then
                                oBPMaster_Target.BPBankAccounts.BPCode = sSynccode 'oBPMaster.BPBankAccounts.BPCode
                            End If
                            oBPMaster_Target.BPBankAccounts.Branch = oBPMaster.BPBankAccounts.Branch
                            oBPMaster_Target.BPBankAccounts.IBAN = oBPMaster.BPBankAccounts.IBAN
                            oBPMaster_Target.BPBankAccounts.AccountName = oBPMaster.BPBankAccounts.AccountName
                            oBPMaster_Target.BPBankAccounts.BICSwiftCode = oBPMaster.BPBankAccounts.BICSwiftCode

                            '*****************ADDED ON 07/09/2015 STARTS**************
                            oBPMaster_Target.BPBankAccounts.MandateID = oBPMaster.BPBankAccounts.MandateID
                            oBPMaster_Target.BPBankAccounts.SignatureDate = oBPMaster.BPBankAccounts.SignatureDate
                            '*****************ADDED ON 07/09/2015 ENDS**************
                        Else
                            For imjd As Integer = 0 To oBPMaster_Target.BPBankAccounts.Count - 1
                                oBPMaster_Target.BPBankAccounts.SetCurrentLine(imjd)
                                ''   MsgBox("target " & oBPMaster_Target.BPBankAccounts.AccountNo & " holding " & oBPMaster.BPBankAccounts.AccountNo & "Internal base " & oBPMaster.BPBankAccounts.InternalKey & " target " & oBPMaster_Target.BPBankAccounts.InternalKey)
                                If Not String.IsNullOrEmpty(oBPMaster.BPBankAccounts.AccountNo) Then
                                    If oBPMaster_Target.BPBankAccounts.AccountNo = oBPMaster.BPBankAccounts.AccountNo Then
                                        oBPMaster_Target.BPBankAccounts.Country = oBPMaster.BPBankAccounts.Country
                                        oBPMaster_Target.BPBankAccounts.AccountNo = oBPMaster.BPBankAccounts.AccountNo
                                        oBPMaster_Target.BPBankAccounts.BankCode = oBPMaster.BPBankAccounts.BankCode
                                        If Not String.IsNullOrEmpty(sSynccode) Then
                                            oBPMaster_Target.BPBankAccounts.BPCode = sSynccode 'oBPMaster.BPBankAccounts.BPCode
                                        End If
                                        oBPMaster_Target.BPBankAccounts.Branch = oBPMaster.BPBankAccounts.Branch
                                        oBPMaster_Target.BPBankAccounts.IBAN = oBPMaster.BPBankAccounts.IBAN
                                        oBPMaster_Target.BPBankAccounts.AccountName = oBPMaster.BPBankAccounts.AccountName
                                        oBPMaster_Target.BPBankAccounts.BICSwiftCode = oBPMaster.BPBankAccounts.BICSwiftCode

                                        '*****************ADDED ON 07/09/2015 STARTS**************
                                        oBPMaster_Target.BPBankAccounts.MandateID = oBPMaster.BPBankAccounts.MandateID
                                        oBPMaster_Target.BPBankAccounts.SignatureDate = oBPMaster.BPBankAccounts.SignatureDate
                                    End If
                                End If

                            Next
                        End If
                    Next
                End If

                '' oBPMaster_Target.SaveXML("E:\Abeo-Projects\PWC\SVN - Copy\1. Source\AE_PWC\Add-ons\Master Data Synchronization\AE_PWC_AO03\bin\Debug\BP.xml")
                ' ''PAYMENT RUN TAB
                oBPMaster_Target.HouseBankCountry = oBPMaster.HouseBankCountry
                oBPMaster_Target.HouseBank = oBPMaster.HouseBank
                oBPMaster_Target.HouseBankAccount = oBPMaster.HouseBankAccount

                ''*****************ADDED ON 07/09/2015 STARTS**************
                oBPMaster_Target.DME = oBPMaster.DME
                oBPMaster_Target.InstructionKey = oBPMaster.InstructionKey
                ''*****************ADDED ON 07/09/2015 ENDS**************
                ' ''ACCOUNTING TAB
                If oBPMaster_Target.AccountRecivablePayables.Count = 0 Then
                    oBPMaster_Target.AccountRecivablePayables.AccountCode = oBPMaster.AccountRecivablePayables.AccountCode
                    oBPMaster_Target.AccountRecivablePayables.Add()
                Else
                    oBPMaster_Target.AccountRecivablePayables.SetCurrentLine(0)
                    oBPMaster_Target.AccountRecivablePayables.AccountCode = oBPMaster.AccountRecivablePayables.AccountCode
                End If

                oBPMaster_Target.VatLiable = oBPMaster.VatLiable
                ' oBPMaster_Target.WithholdingTaxCertified = oBPMaster.WithholdingTaxCertified
                oBPMaster_Target.VatGroup = oBPMaster.VatGroup
                oBPMaster_Target.Frozen = SAPbobsCOM.BoYesNoEnum.tNO

                oBPMaster_Target.FreeText = oBPMaster.FreeText
                oBPMaster_Target.Frozen = oBPMaster.Frozen
                If Not String.IsNullOrEmpty(oBPMaster.UserFields.Fields.Item("U_AB_WTAXREQ").Value) Then
                    oBPMaster_Target.UserFields.Fields.Item("U_AB_WTAXREQ").Value = oBPMaster.UserFields.Fields.Item("U_AB_WTAXREQ").Value
                End If

                If Not String.IsNullOrEmpty(oBPMaster.PeymentMethodCode) Then
                    oBPMaster_Target.PeymentMethodCode = oBPMaster.PeymentMethodCode
                End If
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Catch ex As Exception
                p_oSBOApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with Error " & ex.Message, sFuncName)
            End Try



        End Sub
        Public Sub BP_Assignment_Old(ByRef oBPMaster As SAPbobsCOM.BusinessPartners, ByRef oBPMaster_Target As SAPbobsCOM.BusinessPartners)

            sFuncName = "BP_Assignment()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

            oBPMaster_Target.CardName = oBPMaster.CardName
            oBPMaster_Target.Series = oBPMaster.Series
            oBPMaster_Target.CardType = oBPMaster.CardType
            oBPMaster_Target.GroupCode = oBPMaster.GroupCode
            oBPMaster_Target.Currency = oBPMaster.Currency
            oBPMaster_Target.FederalTaxID = oBPMaster.FederalTaxID
            'GENERAL TAB
            oBPMaster_Target.Phone1 = oBPMaster.Phone1
            oBPMaster_Target.Phone2 = oBPMaster.Phone2
            oBPMaster_Target.Fax = oBPMaster.Fax
            oBPMaster_Target.EmailAddress = oBPMaster.EmailAddress
            oBPMaster_Target.ContactPerson = oBPMaster.ContactPerson
            'CONTACT PERSON  TAB
            If oBPMaster_Target.ContactEmployees.Count = 0 Then
                oBPMaster_Target.ContactEmployees.Add()
                oBPMaster_Target.ContactEmployees.Name = oBPMaster.ContactEmployees.Name
                oBPMaster_Target.ContactEmployees.Position = oBPMaster.ContactEmployees.Position
                oBPMaster_Target.ContactEmployees.E_Mail = oBPMaster.ContactEmployees.E_Mail
                oBPMaster_Target.ContactEmployees.Phone1 = oBPMaster.ContactEmployees.Phone1
                oBPMaster_Target.ContactEmployees.Phone2 = oBPMaster.ContactEmployees.Phone2
            Else
                ' oBPMaster_Target.ContactEmployees.Add()
                oBPMaster_Target.ContactEmployees.Name = oBPMaster.ContactEmployees.Name
                oBPMaster_Target.ContactEmployees.Position = oBPMaster.ContactEmployees.Position
                oBPMaster_Target.ContactEmployees.E_Mail = oBPMaster.ContactEmployees.E_Mail
                oBPMaster_Target.ContactEmployees.Phone1 = oBPMaster.ContactEmployees.Phone1
                oBPMaster_Target.ContactEmployees.Phone2 = oBPMaster.ContactEmployees.Phone2
            End If

            'ADDRESS TAB
            If oBPMaster_Target.Addresses.Count = 0 Then
                For imjs As Integer = 0 To oBPMaster.Addresses.Count - 1
                    oBPMaster_Target.Addresses.AddressName = oBPMaster.Addresses.AddressName
                    oBPMaster_Target.Addresses.AddressName2 = oBPMaster.Addresses.AddressName2
                    oBPMaster_Target.Addresses.AddressName3 = oBPMaster.Addresses.AddressName3
                    oBPMaster_Target.Addresses.Street = oBPMaster.Addresses.Street
                    oBPMaster_Target.Addresses.Block = oBPMaster.Addresses.Block
                    oBPMaster_Target.Addresses.City = oBPMaster.Addresses.City
                    oBPMaster_Target.Addresses.ZipCode = oBPMaster.Addresses.ZipCode
                    oBPMaster_Target.Addresses.Country = oBPMaster.Addresses.Country
                    oBPMaster_Target.Addresses.AddressType = oBPMaster.Addresses.AddressType
                    oBPMaster_Target.Addresses.Add()
                Next imjs
            Else
                For imjs As Integer = 0 To oBPMaster.Addresses.Count - 1
                    oBPMaster.Addresses.SetCurrentLine(imjs)
                    If imjs <= oBPMaster_Target.Addresses.Count - 1 Then
                        oBPMaster_Target.Addresses.SetCurrentLine(imjs)
                        oBPMaster_Target.Addresses.AddressName = oBPMaster.Addresses.AddressName
                        oBPMaster_Target.Addresses.AddressName2 = oBPMaster.Addresses.AddressName2
                        oBPMaster_Target.Addresses.AddressName3 = oBPMaster.Addresses.AddressName3
                        oBPMaster_Target.Addresses.Street = oBPMaster.Addresses.Street
                        oBPMaster_Target.Addresses.Block = oBPMaster.Addresses.Block
                        oBPMaster_Target.Addresses.City = oBPMaster.Addresses.City
                        oBPMaster_Target.Addresses.ZipCode = oBPMaster.Addresses.ZipCode
                        oBPMaster_Target.Addresses.Country = oBPMaster.Addresses.Country
                        oBPMaster_Target.Addresses.AddressType = oBPMaster.Addresses.AddressType
                    Else
                        oBPMaster_Target.Addresses.Add()
                        oBPMaster_Target.Addresses.AddressName = oBPMaster.Addresses.AddressName
                        oBPMaster_Target.Addresses.AddressName2 = oBPMaster.Addresses.AddressName2
                        oBPMaster_Target.Addresses.AddressName3 = oBPMaster.Addresses.AddressName3
                        oBPMaster_Target.Addresses.Street = oBPMaster.Addresses.Street
                        oBPMaster_Target.Addresses.Block = oBPMaster.Addresses.Block
                        oBPMaster_Target.Addresses.City = oBPMaster.Addresses.City
                        oBPMaster_Target.Addresses.ZipCode = oBPMaster.Addresses.ZipCode
                        oBPMaster_Target.Addresses.Country = oBPMaster.Addresses.Country
                        oBPMaster_Target.Addresses.AddressType = oBPMaster.Addresses.AddressType
                    End If
                Next imjs
            End If
            'PAYMENT TERMS TAB
            oBPMaster_Target.PayTermsGrpCode = oBPMaster.PayTermsGrpCode
            oBPMaster_Target.BankCountry = oBPMaster.BankCountry
            oBPMaster_Target.DefaultBankCode = oBPMaster.DefaultBankCode
            MsgBox(oBPMaster.DefaultAccount)
            oBPMaster_Target.DefaultAccount = oBPMaster.DefaultAccount
            oBPMaster_Target.DefaultBranch = oBPMaster.DefaultBranch
            oBPMaster_Target.IBAN = oBPMaster.HouseBankIBAN
            'PAYMENT RUN TAB
            oBPMaster_Target.HouseBankCountry = oBPMaster.HouseBankCountry
            oBPMaster_Target.HouseBank = oBPMaster.HouseBank
            oBPMaster_Target.HouseBankAccount = oBPMaster.HouseBankAccount
            'ACCOUNTING TAB
            ' oBPMaster_Target.AccountRecivablePayables = oBPMaster.AccountRecivablePayables
            oBPMaster_Target.VatLiable = oBPMaster.VatLiable
            ' oBPMaster_Target.WithholdingTaxCertified = oBPMaster.WithholdingTaxCertified
            oBPMaster_Target.VatGroup = oBPMaster.VatGroup
            oBPMaster_Target.Frozen = SAPbobsCOM.BoYesNoEnum.tNO

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

        End Sub

    End Module
End Namespace



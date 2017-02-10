
Imports System.IO


Namespace AE_PWC_AO03
    Module user

        Dim sFuncName As String = String.Empty
        Dim ival As Integer
        Dim IsError As Boolean
        Dim iErr As Integer = 0
        Dim sErr As String = String.Empty
        Dim oMatrix As SAPbouiCOM.Matrix = Nothing



        Public Function UsersSync(ByRef oHoldingCompany As SAPbobsCOM.Company, ByRef oTragetCompany As SAPbobsCOM.Company, ByVal sMasterdatacode As String, _
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
            sFuncName = "UsersSync()"

            Try


                Dim oRset As SAPbobsCOM.Recordset = oHoldingCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim oRset_T As SAPbobsCOM.Recordset = oTragetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                Dim oUser_Holding As SAPbobsCOM.Users = oHoldingCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUsers)
                Dim oUser_Target As SAPbobsCOM.Users = oTragetCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUsers)
                Dim sSQL As String = "SELECT userid FROM OUSR T0 WHERE T0.[USER_CODE] = '" & sMasterdatacode & "'"
                Dim sFileName As String = System.Windows.Forms.Application.StartupPath & "\Users.xml"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("SQL " & sSQL, sFuncName)
                oRset.DoQuery(sSQL)
                Dim iUsercode As Integer = oRset.Fields.Item("userid").Value
                oRset_T.DoQuery(sSQL)
                Dim iUsercode_T As Integer = oRset_T.Fields.Item("userid").Value

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function " & sMasterdatacode, sFuncName)
                If oUser_Holding.GetByKey(iUsercode) Then
                    If File.Exists(sFileName) Then
                        File.Delete(sFileName)
                    End If
                    oHoldingCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
                    oTragetCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
                    oUser_Holding.SaveXML(sFileName)
                    If oUser_Target.GetByKey(iUsercode_T) Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling BP_Assignment()", sFuncName)
                        User_Assignment(oUser_Holding, oUser_Target)
                        ' oUser_Target = oTragetCompany.GetBusinessObjectFromXML(sFileName, 0)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Update the User " & oTragetCompany.CompanyDB, sFuncName)
                        ival = oUser_Target.Update()
                        If ival <> 0 Then
                            IsError = True
                            oTragetCompany.GetLastError(iErr, sErr)
                            Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
                            sErrDesc = sErr
                            UsersSync = RTN_ERROR
                            Exit Function
                        End If
                    Else
                        ''  If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Item_Assignment()", sFuncName)
                        '' oUser_Target.UserCode = oUser_Holding.UserCode
                        ' oUser_Target.UserPassword = oUser_Holding.UserPassword
                        ' User_Assignment(oUser_Holding, oUser_Target)
                        oTragetCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
                        oUser_Target = oTragetCompany.GetBusinessObjectFromXML(sFileName, 0)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add the User " & oTragetCompany.CompanyDB, sFuncName)
                        ival = oUser_Target.Add()
                        If ival <> 0 Then
                            IsError = True
                            oTragetCompany.GetLastError(iErr, sErr)
                            Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
                            sErrDesc = sErr
                            UsersSync = RTN_ERROR
                            Exit Function
                        End If
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS " & sErrDesc, sFuncName)
                Else
                    sErrDesc = "No matching records found in the holding DB " & sMasterdatacode
                    UsersSync = RTN_ERROR
                    Call WriteToLogFile("Completed with ERROR ---" & sErrDesc, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErrDesc, sFuncName)
                    Exit Function
                End If
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS " & sErr, sFuncName)
                UsersSync = RTN_SUCCESS
                sErrDesc = String.Empty
            Catch ex As Exception
                UsersSync = RTN_ERROR
                sErrDesc = ex.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            End Try


        End Function

    End Module
End Namespace


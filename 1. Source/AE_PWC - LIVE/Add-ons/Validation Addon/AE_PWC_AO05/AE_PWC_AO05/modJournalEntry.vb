Module modJournalEntry
    Public dtReverseDate As Date

    Function ReverseDate_Validation(ByRef oForm As SAPbouiCOM.Form, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim oRS As SAPbobsCOM.Recordset
        Dim sQuery As String = String.Empty

        Try
            sFuncName = "POValidation()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            oRS = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            sQuery = " SELECT CONVERT(VARCHAR(25),DATEADD(dd,-(DAY(DATEADD(mm,1,GETDATE()))-1),DATEADD(mm,1,GETDATE())),103) "

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing the Query : " & sQuery, sFuncName)
            oRS.DoQuery(sQuery)

            dtReverseDate = oRS.Fields.Item(0).Value


            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            ReverseDate_Validation = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            ReverseDate_Validation = RTN_ERROR
        End Try
    End Function
End Module

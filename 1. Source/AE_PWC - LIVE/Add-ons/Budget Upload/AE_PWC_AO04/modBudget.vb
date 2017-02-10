
Namespace AE_PWC_AO04

    Module modBudget
        Public Function BudgetSetup_Add(ByRef oform As SAPbouiCOM.Form, ByVal oDT_Budget As DataTable, ByRef sErrDesc As String) As Long

            ' **********************************************************************************
            '   Function    :   BudgetSetup_Add()
            '   Purpose     :   This function will upload the data from CSV file to Dataview
            '   Parameters  :   ByRef CurrFileToUpload AS String 
            '                       CurrFileToUpload = File Name
            '   Author      :   JOHN
            '   Date        :   MAY 2015 27
            ' **********************************************************************************


            Dim sFuncName As String = String.Empty
            Try
                sFuncName = "BudgetSetup_Add()"

                WriteIntoEditBox(oform, "Completed with SUCCESS ..... ", sErrDesc)

            Catch ex As Exception
                sErrDesc = ex.Message
                p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error occured while reading content of  " & ex.Message, sFuncName)
                Call WriteToLogFile(ex.Message, sFuncName)
                Return Nothing
            Finally
            End Try

        End Function

    End Module

End Namespace



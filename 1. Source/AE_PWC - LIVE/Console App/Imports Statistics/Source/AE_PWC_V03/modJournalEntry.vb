Module modJournalEntry


    ''Public Function JournalEntry_Posting(ByVal oDsJournal As DataSet, ByRef oCompany As SAPbobsCOM.Company _
    ''                                     , ByVal sFileName As String, ByRef sErrDesc As String) As Long

    ''    Dim sFuncName As String = String.Empty
    ''    Dim ival As Integer
    ''    Dim IsError As Boolean
    ''    Dim iErr As Integer = 0
    ''    Dim sErr As String = String.Empty
    ''    Dim sJV As String = String.Empty
    ''    Dim dCreditAmount As Double = 0.0
    ''    Dim dDebitAmount As Double = 0.0
    ''    Dim Amount As Double = 0.0
    ''    Dim iseries As Integer = 0
    ''    Dim oRset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    ''    Dim sSQL As String = String.Empty
    ''    Dim sJESeries As String = String.Empty

    ''    Try
    ''        sFuncName = "JournalEntry_Posting"
    ''        Console.WriteLine("Starting Function ", sFuncName)
    ''        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)
    ''        'Fetching Entity from the Dataview object

    ''        sSQL = "SELECT T0.[Series] FROM NNM1 T0 WHERE T0.[ObjectCode] =30 and  T0.[SeriesName] = '" & p_oCompDef.sSeries & "'"
    ''        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Series SQL " & sSQL, sFuncName)
    ''        oRset.DoQuery(sSQL)

    ''        sJESeries = oRset.Fields.Item("Series").Value

    ''        Dim oJournalEntry As SAPbobsCOM.JournalEntries = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)

    ''        ' Journal Entry Document Header Information
    ''        oJournalEntry.ReferenceDate = oDsJournal.Tables(0).Rows(0).Item(4).ToString.Trim
    ''        oJournalEntry.Series = sJESeries
    ''        oJournalEntry.DueDate = oDsJournal.Tables(0).Rows(0).Item(4).ToString.Trim
    ''        oJournalEntry.TaxDate = oDsJournal.Tables(0).Rows(0).Item(4).ToString.Trim
    ''        oJournalEntry.Reference = oDsJournal.Tables(0).Rows(0).Item(3).ToString.Trim
    ''        oJournalEntry.TransactionCode = p_oCompDef.sSeries



    ''        oJournalEntry.Memo = oDsJournal.Tables(0).Rows(0).Item(3).ToString.Trim & " - " & CStr(oDsJournal.Tables(0).Rows(0).Item(4).ToString.Trim)
    ''        oJournalEntry.UserFields.Fields.Item("U_AB_FileName").Value = sFileName

    ''        'Journal Entry Document Line Information
    ''        For mjs As Integer = 0 To oDsJournal.Tables(1).Rows.Count - 1

    ''            oJournalEntry.Lines.AccountCode = oDsJournal.Tables(1).Rows(mjs).Item(0).ToString.Trim
    ''            If CDbl(oDsJournal.Tables(1).Rows(mjs).Item(2).ToString.Trim) > 0 Then

    ''                oJournalEntry.Lines.Debit = Math.Abs(CDbl(oDsJournal.Tables(1).Rows(mjs).Item(2).ToString.Trim))
    ''                oJournalEntry.Lines.Credit = 0
    ''                dDebitAmount += Math.Abs(CDbl(oDsJournal.Tables(1).Rows(mjs).Item(2).ToString.Trim))

    ''            ElseIf CDbl(oDsJournal.Tables(1).Rows(mjs).Item(2).ToString.Trim) < 0 Then

    ''                oJournalEntry.Lines.Debit = 0
    ''                oJournalEntry.Lines.Credit = Math.Abs(CDbl(oDsJournal.Tables(1).Rows(mjs).Item(2).ToString.Trim))
    ''                dCreditAmount += Math.Abs(CDbl(oDsJournal.Tables(1).Rows(mjs).Item(2).ToString.Trim))

    ''            End If
    ''            oJournalEntry.Lines.LineMemo = oDsJournal.Tables(1).Rows(mjs).Item(1).ToString.Trim

    ''            'oJournalEntry.Lines.ReferenceDate1 = drv(1)
    ''            'oJournalEntry.Lines.DueDate = oDsJournal.Tables(1).Rows(mjs).Item(6).ToString.Trim
    ''            'oJournalEntry.Lines.TaxDate = oDsJournal.Tables(1).Rows(mjs).Item(6).ToString.Trim

    ''            'LOS (Dimension 1)
    ''            If Not String.IsNullOrEmpty(oDsJournal.Tables(1).Rows(mjs).Item(9).ToString.Trim) Then
    ''                oJournalEntry.Lines.CostingCode = oDsJournal.Tables(1).Rows(mjs).Item(9).ToString.Trim
    ''            End If

    ''            'BU ( Dimension 2)
    ''            If Not String.IsNullOrEmpty(oDsJournal.Tables(1).Rows(mjs).Item(8).ToString.Trim) Then
    ''                oJournalEntry.Lines.CostingCode2 = oDsJournal.Tables(1).Rows(mjs).Item(8).ToString.Trim
    ''            End If

    ''            'OU ( Dimension 3)
    ''            If Not String.IsNullOrEmpty(oDsJournal.Tables(1).Rows(mjs).Item(3).ToString.Trim) Then
    ''                oJournalEntry.Lines.CostingCode3 = oDsJournal.Tables(1).Rows(mjs).Item(3).ToString.Trim
    ''            End If

    ''            'UDF
    ''            If Not String.IsNullOrEmpty(oDsJournal.Tables(1).Rows(mjs).Item(6).ToString.Trim) Then
    ''                oJournalEntry.Lines.UserFields.Fields.Item("U_AB_PARTNER").Value = oDsJournal.Tables(1).Rows(mjs).Item(6).ToString.Trim
    ''            End If

    ''            oJournalEntry.Lines.Add()
    ''        Next

    ''        Console.WriteLine("Total Credit Amount " & dCreditAmount, sFuncName)
    ''        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Total Credit Amount " & dCreditAmount, sFuncName)
    ''        Console.WriteLine("Total Debit Amount " & dDebitAmount, sFuncName)
    ''        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Total Debit Amount " & dDebitAmount, sFuncName)

    ''        Amount = dCreditAmount - dDebitAmount
    ''        If Amount > 0 Then
    ''            oJournalEntry.Lines.AccountCode = 99999999
    ''            oJournalEntry.Lines.Debit = Amount
    ''            oJournalEntry.Lines.Credit = 0
    ''        Else
    ''            oJournalEntry.Lines.AccountCode = 99999999
    ''            oJournalEntry.Lines.Credit = Amount
    ''            oJournalEntry.Lines.Debit = 0
    ''        End If

    ''        Console.WriteLine("Attempting to Add the Journal Entry ", sFuncName)
    ''        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add the Journal Entry", sFuncName)
    ''        ival = oJournalEntry.Add()

    ''        If ival <> 0 Then
    ''            IsError = True
    ''            oCompany.GetLastError(iErr, sErr)
    ''            sErrDesc = sErr
    ''            Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
    ''            Console.WriteLine("Completed with ERROR " & sErr, sFuncName)
    ''            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
    ''            JournalEntry_Posting = RTN_ERROR
    ''            Exit Function
    ''        End If

    ''        Console.WriteLine("Completed with SUCCESS", sFuncName)
    ''        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS ", sFuncName)
    ''        oCompany.GetNewObjectCode(sJV)
    ''        Console.WriteLine("Journal Entry DocEntry  " & sJV, sFuncName)
    ''        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Journal Entry DocEntry  " & sJV, sFuncName)
    ''        JournalEntry_Posting = RTN_SUCCESS

    ''    Catch ex As Exception

    ''        sErrDesc = ex.Message
    ''        Call WriteToLogFile(sErrDesc, sFuncName)

    ''        Console.WriteLine("Completed with ERROR ", sFuncName)
    ''        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & ex.Message, sFuncName)
    ''        JournalEntry_Posting = RTN_ERROR
    ''        Exit Function
    ''    End Try

    ''End Function

    ''Public Function JournalEntry_Posting(ByVal oDVJournal As DataView, ByRef oCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long

    ''    Dim sFuncName As String = String.Empty
    ''    Dim sEntity As String = String.Empty
    ''    Dim ddate As Date = Nothing
    ''    Dim ival As Integer
    ''    Dim IsError As Boolean
    ''    Dim iErr As Integer = 0
    ''    Dim sErr As String = String.Empty
    ''    Dim sJV As String = String.Empty
    ''    Dim sEmpCat As String = String.Empty
    ''    Dim sPaycode As String = String.Empty
    ''    Dim sCreditGL As String = String.Empty
    ''    Dim sRemarks As String = String.Empty
    ''    Dim sdimension1 As String = String.Empty
    ''    Dim sdimension2 As String = String.Empty
    ''    Dim iindex As Integer = 0
    ''    Dim dCreditAmount As Double = 0.0
    ''    Dim dDebitAmount As Double = 0.0
    ''    Dim Amount As Double = 0.0

    ''    Dim oRecordset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    ''    Dim oProfitcenter As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

    ''    Try
    ''        sFuncName = "JournalEntry_Posting"
    ''        Console.WriteLine("Starting Function ", sFuncName)
    ''        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)
    ''        'Fetching Entity from the Dataview object

    ''        ddate = oDVJournal.Table.Rows(1).Item(2).ToString.Substring(4, 4) & "/" & oDVJournal.Table.Rows(1).Item(2).ToString.Substring(2, 2) & "/" & oDVJournal.Table.Rows(1).Item(2).ToString.Substring(0, 2)
    ''        Dim oJournalEntry As SAPbobsCOM.JournalEntries = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
    ''        ' Journal Entry Document Header Information
    ''        oJournalEntry.ReferenceDate = ddate
    ''        oJournalEntry.DueDate = ddate
    ''        oJournalEntry.TaxDate = ddate
    ''        oJournalEntry.Reference = oDVJournal.Table.Rows(1).Item(1).ToString.Trim
    ''        oJournalEntry.Memo = oDVJournal.Table.Rows(1).Item(5).ToString.Trim
    ''        'Journal Entry Document Line Information
    ''        For Each drv As DataRowView In oDVJournal

    ''            '-------------------------------------------------------------------------------
    ''            '------- Checking sCreditGL is empty or not, First time this Credit GL varaibles is empty, so it will satisfied the condition.

    ''            oProfitcenter.DoQuery("SELECT T0.[GrpCode] FROM OPRC T0 WHERE T0.[PrcCode]  = '" & drv(13).ToString.Trim & "' and  T0.[DimCode] = '" & p_oCompDef.ProjectDimension & "'")

    ''            If String.IsNullOrEmpty(sCreditGL) Then

    ''                'Account Code
    ''                ' Dim dd = "SELECT T0.[U_AE_GLAcct] FROM [dbo].[@AE_EMPCATGL]  T0 WHERE T0.[U_AE_Category] = '" & drv(6).ToString.Trim & "' and   T0.[U_AE_PayItem] = '" & drv(7).ToString.Trim & "'"
    ''                oRecordset.DoQuery("SELECT T0.[U_AE_GLAcct] FROM [dbo].[@AE_EMPCATGL]  T0 WHERE T0.[U_AE_Category] = '" & drv(6).ToString.Trim & "' and   T0.[U_AE_PayItem] = '" & drv(7).ToString.Trim & "'")
    ''                If Not String.IsNullOrEmpty(oRecordset.Fields.Item("U_AE_GLAcct").Value) Then
    ''                    oJournalEntry.Lines.AccountCode = oRecordset.Fields.Item("U_AE_GLAcct").Value.ToString.Trim
    ''                Else
    ''                    oJournalEntry.Lines.AccountCode = drv(8).ToString.Trim
    ''                End If
    ''                ' Debit Amount
    ''                If Not String.IsNullOrEmpty(drv(9).ToString.Trim) Then
    ''                    oJournalEntry.Lines.Debit = CDbl(drv(9).ToString.Trim)
    ''                    oJournalEntry.Lines.Credit = "0"
    ''                End If
    ''                'Line Remarks
    ''                oJournalEntry.Lines.LineMemo = drv(12).ToString.Trim
    ''                'Line Project Code ( Dimension 1)
    ''                If Not String.IsNullOrEmpty(drv(13).ToString.Trim) Then
    ''                    oJournalEntry.Lines.CostingCode = drv(13).ToString.Trim
    ''                End If

    ''                'Line Profit Center ( Dimension 2)
    ''                If Not String.IsNullOrEmpty(oProfitcenter.Fields.Item("GrpCode").Value) Then
    ''                    oJournalEntry.Lines.CostingCode2 = oProfitcenter.Fields.Item("GrpCode").Value
    ''                End If

    ''                'Line Cost Center ( Dimension 3)
    ''                If Not String.IsNullOrEmpty(drv(14).ToString.Trim) Then
    ''                    oJournalEntry.Lines.CostingCode3 = drv(14).ToString.Trim
    ''                End If
    ''                oJournalEntry.Lines.Add()

    ''            Else
    ''                '-----------------------------------------------------------------------------------------
    ''                '-------- Here we check the sCreditGL is equal to CreditGL in the dataview ---------------
    ''                '-------- Its equal True part will executed otherwise Else part will execute -------------
    ''                If sCreditGL = drv(10).ToString.Trim Then

    ''                    '----------------------- sCreditGL and Credit GL in Dataview is equal we just accumulate the 
    ''                    '                        Credit amount in the bCreditAmount Variable 
    ''                    'Account Code
    ''                    ' Dim dd = "SELECT T0.[U_AE_GLAcct] FROM [dbo].[@AE_EMPCATGL]  T0 WHERE T0.[U_AE_Category] = '" & drv(6).ToString.Trim & "' and   T0.[U_AE_PayItem] = '" & drv(7).ToString.Trim & "'"
    ''                    oRecordset.DoQuery("SELECT T0.[U_AE_GLAcct] FROM [dbo].[@AE_EMPCATGL]  T0 WHERE T0.[U_AE_Category] = '" & drv(6).ToString.Trim & "' and   T0.[U_AE_PayItem] = '" & drv(7).ToString.Trim & "'")
    ''                    If Not String.IsNullOrEmpty(oRecordset.Fields.Item("U_AE_GLAcct").Value) Then
    ''                        oJournalEntry.Lines.AccountCode = oRecordset.Fields.Item("U_AE_GLAcct").Value.ToString.Trim
    ''                    Else
    ''                        oJournalEntry.Lines.AccountCode = drv(8).ToString.Trim
    ''                    End If
    ''                    ' Debit Amount
    ''                    If Not String.IsNullOrEmpty(drv(9).ToString.Trim) Then
    ''                        oJournalEntry.Lines.Debit = CDbl(drv(9).ToString.Trim)
    ''                        oJournalEntry.Lines.Credit = "0"
    ''                    End If
    ''                    'Line Remarks
    ''                    oJournalEntry.Lines.LineMemo = drv(12).ToString.Trim
    ''                    'Line Project Code ( Dimension 1)
    ''                    If Not String.IsNullOrEmpty(drv(13).ToString.Trim) Then
    ''                        oJournalEntry.Lines.CostingCode = drv(13).ToString.Trim
    ''                    End If

    ''                    'Line Profit Center ( Dimension 2)
    ''                    If Not String.IsNullOrEmpty(oProfitcenter.Fields.Item("GrpCode").Value) Then
    ''                        oJournalEntry.Lines.CostingCode2 = oProfitcenter.Fields.Item("GrpCode").Value
    ''                    End If

    ''                    'Line Cost Center ( Dimension 2)
    ''                    If Not String.IsNullOrEmpty(drv(14).ToString.Trim) Then
    ''                        oJournalEntry.Lines.CostingCode3 = drv(14).ToString.Trim
    ''                    End If
    ''                    oJournalEntry.Lines.Add()

    ''                ElseIf sCreditGL <> drv(10).ToString.Trim Then

    ''                    '----------------------- sCreditGL and Credit GL in Dataview is not equal
    ''                    '                        We adding Credit GL, Credit Amount, Remarks, Dimentions and reset the Credit amount to zero

    ''                    'Account Code
    ''                    oJournalEntry.Lines.AccountCode = sCreditGL
    ''                    ' Credit Amount
    ''                    oJournalEntry.Lines.Debit = 0
    ''                    oJournalEntry.Lines.Credit = dCreditAmount
    ''                    'Line Remarks
    ''                    oJournalEntry.Lines.LineMemo = sRemarks
    ''                    oJournalEntry.Lines.Add()

    ''                    dCreditAmount = 0.0
    ''                    '----------------------- Adding information regarding Debit side

    ''                    'Account Code
    ''                    oRecordset.DoQuery("SELECT T0.[U_AE_GLAcct] FROM [dbo].[@AE_EMPCATGL]  T0 WHERE T0.[U_AE_Category] = '" & drv(6).ToString.Trim & "' and   T0.[U_AE_PayItem] = '" & drv(7).ToString.Trim & "'")
    ''                    If Not String.IsNullOrEmpty(oRecordset.Fields.Item("U_AE_GLAcct").Value) Then
    ''                        oJournalEntry.Lines.AccountCode = oRecordset.Fields.Item("U_AE_GLAcct").Value.ToString.Trim
    ''                    Else
    ''                        oJournalEntry.Lines.AccountCode = drv(8).ToString.Trim
    ''                    End If
    ''                    ' Debit Amount
    ''                    If Not String.IsNullOrEmpty(drv(9).ToString.Trim) Then
    ''                        oJournalEntry.Lines.Debit = CDbl(drv(9).ToString.Trim)
    ''                        oJournalEntry.Lines.Credit = "0"
    ''                    End If
    ''                    'Line Remarks
    ''                    oJournalEntry.Lines.LineMemo = drv(12).ToString.Trim
    ''                    'Line Project Code ( Dimension 1)
    ''                    If Not String.IsNullOrEmpty(drv(13).ToString.Trim) Then
    ''                        oJournalEntry.Lines.CostingCode = drv(13).ToString.Trim
    ''                    End If

    ''                    'Line Profit Center ( Dimension 2)
    ''                    If Not String.IsNullOrEmpty(oProfitcenter.Fields.Item("GrpCode").Value) Then
    ''                        oJournalEntry.Lines.CostingCode2 = oProfitcenter.Fields.Item("GrpCode").Value
    ''                    End If

    ''                    'Line Cost Center ( Dimension 2)
    ''                    If Not String.IsNullOrEmpty(drv(14).ToString.Trim) Then
    ''                        oJournalEntry.Lines.CostingCode3 = drv(14).ToString.Trim
    ''                    End If
    ''                    oJournalEntry.Lines.Add()

    ''                End If

    ''            End If
    ''            sEmpCat = drv(6).ToString.Trim
    ''            sPaycode = drv(7).ToString.Trim
    ''            iindex += 1
    ''            dCreditAmount += drv(11).ToString.Trim
    ''            sCreditGL = drv(10).ToString.Trim
    ''            sRemarks = drv(12).ToString.Trim
    ''            sdimension1 = drv(13).ToString.Trim
    ''            sdimension2 = drv(14).ToString.Trim

    ''        Next

    ''        'Account Code
    ''        oJournalEntry.Lines.AccountCode = sCreditGL
    ''        ' Credit Amount
    ''        oJournalEntry.Lines.Debit = 0
    ''        oJournalEntry.Lines.Credit = dCreditAmount
    ''        'Line Remarks
    ''        oJournalEntry.Lines.LineMemo = sRemarks


    ''        Console.WriteLine("Attempting to Add the Journal Entry ", sFuncName)

    ''        Console.WriteLine("Total amount  ", dCreditAmount)
    ''        WriteToLogFile_Debug("Total Amount ---", dCreditAmount)
    ''        dCreditAmount = 0.0

    ''        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add the Journal Entry", sFuncName)
    ''        ival = oJournalEntry.Add()

    ''        If ival <> 0 Then
    ''            IsError = True
    ''            oCompany.GetLastError(iErr, sErr)
    ''            p_sJournalEntryError = "Error Code :- " & iErr & " Error Description :- " & sErr
    ''            Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
    ''            Console.WriteLine("Completed with ERROR ", sFuncName)
    ''            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
    ''            JournalEntry_Posting = RTN_ERROR
    ''            Exit Function
    ''        End If

    ''        Console.WriteLine("Completed with SUCCESS", sFuncName)
    ''        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS ", sFuncName)
    ''        oCompany.GetNewObjectCode(sJV)
    ''        Console.WriteLine("Journal Entry DocEntry  " & sJV, sFuncName)
    ''        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Journal Entry DocEntry  " & sJV, sFuncName)
    ''        JournalEntry_Posting = RTN_SUCCESS

    ''    Catch ex As Exception
    ''        p_sJournalEntryError = "Error Description :- " & ex.Message
    ''        Call WriteToLogFile(ex.Message, sFuncName)
    ''        Console.WriteLine("Completed with ERROR ", sFuncName)
    ''        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & ex.Message, sFuncName)
    ''        JournalEntry_Posting = RTN_ERROR
    ''        Exit Function
    ''    End Try

    ''End Function

End Module

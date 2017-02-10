Imports System.Globalization

Module modJournalEntry


    Public Function JournalEntry_Posting_Testing(ByVal oDVJour As DataView, ByRef oCompany As SAPbobsCOM.Company, ByVal sFileName As String, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim ival As Integer
        Dim IsError As Boolean
        Dim iErr As Integer = 0
        Dim sErr As String = String.Empty
        Dim sJV As String = String.Empty
        Dim dCreditAmount As Decimal = 0.0
        Dim dDebitAmount As Decimal = 0.0
        Dim Amount As Double = 0.0
        Dim iseries As Integer = 0
        Dim sIpowerPeriod() As String
        Dim iMonthNumber As Integer
        Dim iYear As Integer
        Dim icount As Integer = 1
        Dim oDVJournal As DataView
        Dim oDTAutoIncrement As New DataTable
        Dim iStart As Integer = 1
        Dim iEnd As Integer = 0
        Dim ioDVCount As Integer = 0
        Dim iLoop As Integer = 0

        Dim oRset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim sSQL As String = String.Empty
        Dim sPostingDate As String = String.Empty

        Try
            sFuncName = "JournalEntry_Posting"
            Console.WriteLine("Starting Function ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)
            Dim oJournalEntry As SAPbobsCOM.JournalVouchers = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalVouchers)

            oDTAutoIncrement = oDVJour.ToTable
            oDTAutoIncrement = MergeAutoNumberedToDataTable(oDTAutoIncrement, sErrDesc)
            oDVJournal = New DataView(oDTAutoIncrement)

            'Fetching Entity from the Dataview object
            sSQL = "SELECT T0.[Series] FROM NNM1 T0 WHERE T0.[ObjectCode] =30 and  T0.[SeriesName] = '" & p_oCompDef.sSeries & "'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Series SQL " & sSQL, sFuncName)
            oRset.DoQuery(sSQL)
            iseries = oRset.Fields.Item("Series").Value

            sSQL = "SELECT T0.[Name] FROM [dbo].[@AB_IPOWERPERIOD]  T0 WHERE T0.[Code] = '" & oDVJournal.Table.Rows(0).Item("Code").ToString.Trim & "'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("AB_IPOWERPERIOD  SQL " & sSQL, sFuncName)
            oRset.DoQuery(sSQL)
            sIpowerPeriod = Split(oRset.Fields.Item("Name").Value, " ", True)
            iMonthNumber = DateTime.ParseExact(sIpowerPeriod(0), "MMMM", CultureInfo.CurrentCulture).Month

            If iMonthNumber >= 1 And iMonthNumber <= Month(Now.Date) Then
                iYear = CInt(oDVJournal.Table.Rows(0).Item("Year").ToString.Trim)
            Else
                iYear = CInt(oDVJournal.Table.Rows(0).Item("Year").ToString.Trim) - 1
            End If

            sPostingDate = CStr(iYear) & iMonthNumber.ToString.PadLeft(2, "0"c) & sIpowerPeriod(1)

            ioDVCount = oDVJournal.Count
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(" Total No. of Rows " & ioDVCount, sFuncName)

            If ioDVCount >= 50001 Then
                ioDVCount = ioDVCount / 5
                iStart = 1
                iEnd = ioDVCount
                iLoop = 5
            ElseIf ioDVCount >= 30001 And ioDVCount <= 50000 Then
                ioDVCount = ioDVCount / 3
                iStart = 1
                iEnd = ioDVCount
                iLoop = 3
            ElseIf ioDVCount >= 15000 And ioDVCount <= 30000 Then
                ioDVCount = ioDVCount / 2
                iStart = 1
                iEnd = ioDVCount
                iLoop = 2
            Else
                iStart = 1
                iEnd = ioDVCount
                iLoop = 1
            End If


            For imjs As Integer = iStart To iLoop
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(" Attempting Journal Entry Line " & iStart & " to " & iEnd, sFuncName)
                dCreditAmount = 0.0
                dDebitAmount = 0.0
                Amount = 0.0
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Credit Amount " & CStr(dCreditAmount) & " Debit Amount " & CStr(dDebitAmount) & " Amount " & CStr(Amount), sFuncName)
                oDVJournal.RowFilter = "SNo>=" & iStart & " and SNo<=" & iEnd & ""
                oJournalEntry.JournalEntries.ReferenceDate = DateTime.ParseExact(sPostingDate, "yyyyMMdd", Nothing) ''oDVJournal.Table.Rows(0).Item(1).ToString.Trim
                oJournalEntry.JournalEntries.Series = iseries
                oJournalEntry.JournalEntries.UserFields.Fields.Item("U_AB_FileName").Value = sFileName
                Console.WriteLine(" Attempting Journal Entry Line " & iStart & " to " & iEnd, sFuncName)
                For Each drv As DataRowView In oDVJournal
                    Console.WriteLine(" Adding Line information " & iStart & " / " & iEnd, sFuncName)
                    oJournalEntry.JournalEntries.Lines.AccountCode = drv(0).ToString.Trim
                    If Right(drv(3).ToString.Trim.ToUpper, 1) = "D" Then
                        oJournalEntry.JournalEntries.Lines.Debit = CDbl(Left(drv(3).ToString.Trim, Len(drv(3).ToString.Trim) - 1))
                        oJournalEntry.JournalEntries.Lines.Credit = 0
                        dDebitAmount += CDbl(Left(drv(3).ToString.Trim, Len(drv(3).ToString.Trim) - 1))
                    ElseIf Right(drv(3).ToString.Trim.ToUpper, 1) = "C" Then
                        oJournalEntry.JournalEntries.Lines.Debit = 0
                        oJournalEntry.JournalEntries.Lines.Credit = CDbl(Left(drv(3).ToString.Trim, Len(drv(3).ToString.Trim) - 1))
                        dCreditAmount += CDbl(Left(drv(3).ToString.Trim, Len(drv(3).ToString.Trim) - 1))
                    End If
                    oJournalEntry.JournalEntries.Lines.LineMemo = Left(drv(7).ToString.Trim, 45)
                    oJournalEntry.JournalEntries.Lines.TaxDate = drv(1)
                    oJournalEntry.JournalEntries.Lines.Reference1 = drv(4).ToString.Trim & " - " & drv(5).ToString.Trim 'drv(8).ToString.Trim & "-" & drv(7).ToString.Trim
                    oJournalEntry.JournalEntries.Lines.Reference2 = drv(15).ToString.Trim

                    ' LOS (Dimension 1)
                    If Not String.IsNullOrEmpty(drv(18).ToString.Trim) Then
                        oJournalEntry.JournalEntries.Lines.CostingCode = drv(18).ToString.Trim
                    End If

                    'BU ( Dimension 2)
                    If Not String.IsNullOrEmpty(drv(17).ToString.Trim) Then
                        oJournalEntry.JournalEntries.Lines.CostingCode2 = drv(17).ToString.Trim
                    End If

                    'OU ( Dimension 3)
                    If Not String.IsNullOrEmpty(drv(10).ToString.Trim) Then
                        oJournalEntry.JournalEntries.Lines.CostingCode3 = drv(10).ToString.Trim
                    End If
                    oJournalEntry.JournalEntries.Lines.Add()

                    '  oJournalEntry.Lines.SetCurrentLine(imjs)
                    ' If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Line information " & icount & " / " & oDVJournal.Count & "  " & oCompany.CompanyDB, sFuncName)
                    ' icount += 1
                    iStart += 1
                Next

                Amount = dCreditAmount - dDebitAmount

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Credit Amount " & CStr(dCreditAmount) & " Debit Amount " & CStr(dDebitAmount) & " Amount " & CStr(Amount), sFuncName)
                If Amount < 0 Then
                    oJournalEntry.JournalEntries.Lines.AccountCode = "21161300"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(" Unbalanced Amount Credit Side " & Str(Amount), sFuncName)
                    oJournalEntry.JournalEntries.Lines.Credit = Amount * -1
                    oJournalEntry.JournalEntries.Lines.Debit = 0
                    ' LOS (Dimension 1)
                    oJournalEntry.JournalEntries.Lines.CostingCode = "ZZZZZ2"

                    'BU ( Dimension 2)
                    oJournalEntry.JournalEntries.Lines.CostingCode2 = "ZZZZZ1"

                    'OU ( Dimension 3)
                    oJournalEntry.JournalEntries.Lines.CostingCode3 = "ZZZZZ0"
                ElseIf Amount > 0 Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(" Unbalanced Amount Debit Side " & Str(Amount), sFuncName)
                    oJournalEntry.JournalEntries.Lines.AccountCode = "21161300"
                    oJournalEntry.JournalEntries.Lines.Credit = 0
                    oJournalEntry.JournalEntries.Lines.Debit = Amount
                    ' LOS (Dimension 1)
                    oJournalEntry.JournalEntries.Lines.CostingCode = "ZZZZZ2"

                    'BU ( Dimension 2)
                    oJournalEntry.JournalEntries.Lines.CostingCode2 = "ZZZZZ1"

                    'OU ( Dimension 3)
                    oJournalEntry.JournalEntries.Lines.CostingCode3 = "ZZZZZ0"

                End If

                Console.WriteLine("Attempting to Add the Journal Entry ", sFuncName)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add the Journal Entry", sFuncName)
                'oCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
                ' oJournalEntry.SaveXML(p_oCompDef.sLogPath & "\JE1.xml")
                ival = oJournalEntry.Add()

                If ival <> 0 Then
                    IsError = True
                    oCompany.GetLastError(iErr, sErr)
                    Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
                    Console.WriteLine("Completed with ERROR ", sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
                    JournalEntry_Posting_Testing = RTN_ERROR
                    Exit Function
                End If

                Console.WriteLine("Completed with SUCCESS", sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS ", sFuncName)
                oCompany.GetNewObjectCode(sJV)
                Console.WriteLine("Journal Entry DocEntry  " & sJV, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Journal Entry DocEntry  " & sJV, sFuncName)

                iStart = iEnd + 1
                iEnd = iStart + ioDVCount
            Next imjs
            JournalEntry_Posting_Testing = RTN_SUCCESS
            ' Journal Entry Document Header Information
            ' oJournalEntry.DueDate = DateTime.ParseExact(sPostingDate, "yyyyMMdd", Nothing) ''oDVJournal.Table.Rows(0).Item(1).ToString.Trim
            ' oJournalEntry.TaxDate = DateTime.ParseExact(sPostingDate, "yyyyMMdd", Nothing) ''oDVJournal.Table.Rows(0).Item(1).ToString.Trim
            'oJournalEntry.Reference2 = oDVJournal.Table.Rows(0).Item(15).ToString.Trim
            'oJournalEntry.UserFields.Fields.Item("U_AB_FileName").Value = oDVJournal.Table.Rows(0).Item(16).ToString.Trim
            'Journal Entry Document Line Information
        Catch ex As Exception

            Call WriteToLogFile(ex.Message, sFuncName)
            Console.WriteLine("Completed with ERROR ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & ex.Message, sFuncName)
            JournalEntry_Posting_Testing = RTN_ERROR
            Exit Function
        End Try

    End Function


    Public Function JournalEntry_Posting(ByVal oDVJour As DataView, ByRef oCompany As SAPbobsCOM.Company, ByVal sFileName As String, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim ival As Integer
        Dim IsError As Boolean
        Dim iErr As Integer = 0
        Dim sErr As String = String.Empty
        Dim sJV As String = String.Empty
        Dim dCreditAmount As Decimal = 0.0
        Dim dDebitAmount As Decimal = 0.0
        Dim Amount As Double = 0.0
        Dim iseries As Integer = 0
        Dim sIpowerPeriod() As String
        Dim iMonthNumber As Integer
        Dim iYear As Integer
        Dim icount As Integer = 1
        Dim oDVJournal As DataView
        Dim oDTAutoIncrement As New DataTable
        Dim iStart As Integer = 1
        Dim iEnd As Integer = 0
        Dim ioDVCount As Integer = 0
        Dim iLoop As Integer = 0

        Dim oRset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim sSQL As String = String.Empty
        Dim sPostingDate As String = String.Empty

        Try
            sFuncName = "JournalEntry_Posting"
            Console.WriteLine("Starting Function ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

            oDTAutoIncrement = oDVJour.ToTable
            oDTAutoIncrement = MergeAutoNumberedToDataTable(oDTAutoIncrement, sErrDesc)
            oDVJournal = New DataView(oDTAutoIncrement)

            'Fetching Entity from the Dataview object
            sSQL = "SELECT T0.[Series] FROM NNM1 T0 WHERE T0.[ObjectCode] =30 and  T0.[SeriesName] = '" & p_oCompDef.sSeries & "'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Series SQL " & sSQL, sFuncName)
            oRset.DoQuery(sSQL)
            iseries = oRset.Fields.Item("Series").Value

            sSQL = "SELECT T0.[Name] FROM [dbo].[@AB_IPOWERPERIOD]  T0 WHERE T0.[Code] = '" & oDVJournal.Table.Rows(0).Item("Code").ToString.Trim & "'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("AB_IPOWERPERIOD  SQL " & sSQL, sFuncName)
            oRset.DoQuery(sSQL)
            sIpowerPeriod = Split(oRset.Fields.Item("Name").Value, " ", True)
            iMonthNumber = DateTime.ParseExact(sIpowerPeriod(0), "MMMM", CultureInfo.CurrentCulture).Month

            If iMonthNumber >= 1 And iMonthNumber <= Month(Now.Date) And Year(Now.Date) = oDVJournal.Table.Rows(0).Item("Year").ToString.Trim Then
                iYear = CInt(oDVJournal.Table.Rows(0).Item("Year").ToString.Trim)
            ElseIf Year(Now.Date) = oDVJournal.Table.Rows(0).Item("Year").ToString.Trim Then
                iYear = CInt(oDVJournal.Table.Rows(0).Item("Year").ToString.Trim) - 1
            Else
                iYear = CInt(oDVJournal.Table.Rows(0).Item("Year").ToString.Trim)
            End If

            sPostingDate = CStr(iYear) & iMonthNumber.ToString.PadLeft(2, "0"c) & sIpowerPeriod(1)

            ioDVCount = oDVJournal.Count
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(" Total No. of Rows " & ioDVCount, sFuncName)

            If ioDVCount >= 50001 Then
                ioDVCount = ioDVCount / 5
                iStart = 1
                iEnd = ioDVCount
                iLoop = 5
            ElseIf ioDVCount >= 30001 And ioDVCount <= 50000 Then
                ioDVCount = ioDVCount / 3
                iStart = 1
                iEnd = ioDVCount
                iLoop = 3
            ElseIf ioDVCount >= 15000 And ioDVCount <= 30000 Then
                ioDVCount = ioDVCount / 2
                iStart = 1
                iEnd = ioDVCount
                iLoop = 2
            Else
                iStart = 1
                iEnd = ioDVCount
                iLoop = 1
            End If


            For imjs As Integer = iStart To iLoop
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(" Attempting Journal Entry Line " & iStart & " to " & iEnd, sFuncName)
                dCreditAmount = 0.0
                dDebitAmount = 0.0
                Amount = 0.0
                Dim oJournalEntry As SAPbobsCOM.JournalEntries = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Credit Amount " & CStr(dCreditAmount) & " Debit Amount " & CStr(dDebitAmount) & " Amount " & CStr(Amount), sFuncName)
                oDVJournal.RowFilter = "SNo>=" & iStart & " and SNo<=" & iEnd & ""
                oJournalEntry.ReferenceDate = DateTime.ParseExact(sPostingDate, "yyyyMMdd", Nothing) ''oDVJournal.Table.Rows(0).Item(1).ToString.Trim
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Reference Date  " & DateTime.ParseExact(sPostingDate, "yyyyMMdd", Nothing), sFuncName)
                oJournalEntry.Series = iseries
                oJournalEntry.UserFields.Fields.Item("U_AB_FileName").Value = sFileName
                Console.WriteLine(" Attempting Journal Entry Line " & iStart & " to " & iEnd, sFuncName)
                For Each drv As DataRowView In oDVJournal
                    Console.WriteLine(" Adding Line information " & iStart & " / " & iEnd, sFuncName)
                    oJournalEntry.Lines.AccountCode = drv(0).ToString.Trim
                    If Right(drv(3).ToString.Trim.ToUpper, 1) = "D" Then
                        oJournalEntry.Lines.Debit = CDbl(Left(drv(3).ToString.Trim, Len(drv(3).ToString.Trim) - 1))
                        oJournalEntry.Lines.Credit = 0
                        dDebitAmount += CDbl(Left(drv(3).ToString.Trim, Len(drv(3).ToString.Trim) - 1))
                    ElseIf Right(drv(3).ToString.Trim.ToUpper, 1) = "C" Then
                        oJournalEntry.Lines.Debit = 0
                        oJournalEntry.Lines.Credit = CDbl(Left(drv(3).ToString.Trim, Len(drv(3).ToString.Trim) - 1))
                        dCreditAmount += CDbl(Left(drv(3).ToString.Trim, Len(drv(3).ToString.Trim) - 1))
                    End If
                    oJournalEntry.Lines.LineMemo = Left(drv(7).ToString.Trim, 45)
                    oJournalEntry.Lines.TaxDate = drv(1)
                    '' If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Tax Date  " & Format(drv(1), "yyyyMMdd"), sFuncName)
                    oJournalEntry.Lines.Reference1 = drv(4).ToString.Trim & " - " & drv(5).ToString.Trim 'drv(8).ToString.Trim & "-" & drv(7).ToString.Trim
                    oJournalEntry.Lines.Reference2 = drv(15).ToString.Trim

                    ' LOS (Dimension 1)
                    If Not String.IsNullOrEmpty(drv(18).ToString.Trim) Then
                        oJournalEntry.Lines.CostingCode = drv(18).ToString.Trim
                    End If

                    'BU ( Dimension 2)
                    If Not String.IsNullOrEmpty(drv(17).ToString.Trim) Then
                        oJournalEntry.Lines.CostingCode2 = drv(17).ToString.Trim
                    End If

                    'OU ( Dimension 3)
                    If Not String.IsNullOrEmpty(drv(10).ToString.Trim) Then
                        oJournalEntry.Lines.CostingCode3 = drv(10).ToString.Trim
                    End If
                    oJournalEntry.Lines.Add()

                    '  oJournalEntry.Lines.SetCurrentLine(imjs)
                    ' If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Line information " & icount & " / " & oDVJournal.Count & "  " & oCompany.CompanyDB, sFuncName)
                    ' icount += 1
                    iStart += 1
                Next

                Amount = dCreditAmount - dDebitAmount

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Credit Amount " & CStr(dCreditAmount) & " Debit Amount " & CStr(dDebitAmount) & " Amount " & CStr(Amount), sFuncName)
                If Amount < 0 Then
                    oJournalEntry.Lines.AccountCode = "21161300"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(" Unbalanced Amount Credit Side " & Str(Amount), sFuncName)
                    oJournalEntry.Lines.Credit = Amount * -1
                    oJournalEntry.Lines.Debit = 0
                    ' LOS (Dimension 1)
                    oJournalEntry.Lines.CostingCode = "ZZZZZ2"

                    'BU ( Dimension 2)
                    oJournalEntry.Lines.CostingCode2 = "ZZZZZ1"

                    'OU ( Dimension 3)
                    oJournalEntry.Lines.CostingCode3 = "ZZZZZ0"
                ElseIf Amount > 0 Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(" Unbalanced Amount Debit Side " & Str(Amount), sFuncName)
                    oJournalEntry.Lines.AccountCode = "21161300"
                    oJournalEntry.Lines.Credit = 0
                    oJournalEntry.Lines.Debit = Amount
                    ' LOS (Dimension 1)
                    oJournalEntry.Lines.CostingCode = "ZZZZZ2"

                    'BU ( Dimension 2)
                    oJournalEntry.Lines.CostingCode2 = "ZZZZZ1"

                    'OU ( Dimension 3)
                    oJournalEntry.Lines.CostingCode3 = "ZZZZZ0"

                End If

                Console.WriteLine("Attempting to Add the Journal Entry ", sFuncName)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add the Journal Entry", sFuncName)
                'oCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
                'oJournalEntry.SaveXML(p_oCompDef.sLogPath & "\JE1.xml")
                ival = oJournalEntry.Add()

                If ival <> 0 Then
                    IsError = True
                    oCompany.GetLastError(iErr, sErr)
                    Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
                    Console.WriteLine("Completed with ERROR ", sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
                    JournalEntry_Posting = RTN_ERROR
                    Exit Function
                End If

                Console.WriteLine("Completed with SUCCESS", sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS ", sFuncName)
                oCompany.GetNewObjectCode(sJV)
                Console.WriteLine("Journal Entry DocEntry  " & sJV, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Journal Entry DocEntry  " & sJV, sFuncName)

                iStart = iEnd + 1
                iEnd = iStart + ioDVCount
            Next imjs
            JournalEntry_Posting = RTN_SUCCESS
            ' Journal Entry Document Header Information
            ' oJournalEntry.DueDate = DateTime.ParseExact(sPostingDate, "yyyyMMdd", Nothing) ''oDVJournal.Table.Rows(0).Item(1).ToString.Trim
            ' oJournalEntry.TaxDate = DateTime.ParseExact(sPostingDate, "yyyyMMdd", Nothing) ''oDVJournal.Table.Rows(0).Item(1).ToString.Trim
            'oJournalEntry.Reference2 = oDVJournal.Table.Rows(0).Item(15).ToString.Trim
            'oJournalEntry.UserFields.Fields.Item("U_AB_FileName").Value = oDVJournal.Table.Rows(0).Item(16).ToString.Trim
            'Journal Entry Document Line Information
        Catch ex As Exception

            Call WriteToLogFile(ex.Message, sFuncName)
            Console.WriteLine("Completed with ERROR ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & ex.Message, sFuncName)
            JournalEntry_Posting = RTN_ERROR
            Exit Function
        End Try

    End Function

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

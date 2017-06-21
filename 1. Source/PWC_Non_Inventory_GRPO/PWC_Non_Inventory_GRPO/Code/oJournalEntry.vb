Imports System.Data.SqlClient
Imports System.Globalization
Public Class oJournalEntry

    Public Sub Insert_JE(ByVal ConnectionString As String, ByVal DBName As String)
        Try
            Dim connection As SqlConnection
            connection = New SqlConnection(ConnectionString)
            connection.Open()
            Dim cn As New Connection
            Dim Str As String = "select ID from [AB_GRPO_NON_INV] with (nolock) where Month(SendDate)= Month(getdate()) and  Year(SendDate)= Year(getdate())"
            Dim dt As DataTable = cn.Integration_RunQuery_BR(Str, DBName)
            If dt.Rows.Count = 0 Then
                Dim da As New SqlDataAdapter("EXEC GRPO_NON_INV", ConnectionString)
                Dim DtSet As New System.Data.DataSet
                da.Fill(DtSet)
                connection.Close()
                Dim rd As DataTableReader = DtSet.Tables(0).CreateDataReader()
                connection.Open()
                Using copy As New SqlBulkCopy(connection)
                    copy.DestinationTableName = "AB_GRPO_NON_INV"
                    copy.WriteToServer(rd)
                End Using
            End If
            connection.Close()
        Catch ex As Exception
            Functions.WriteLog("Insert_JE Error Msg:" & ex.Message)

        End Try

    End Sub

    Public Sub CreateJE_LastMonth(ByVal ConnectionString As String, ByVal DBName As String)
        Dim oJE As SAPbobsCOM.JournalEntries
        Dim sqlConx As SqlConnection = New SqlConnection(ConnectionString)
        Try
            Dim cn As New Connection
            Dim xm As New oXML


            Dim oCompany As SAPbobsCOM.Company = PublicVariable.oCompanyInfo
            Dim query As String
            query = "SELECT [ID],[DocEntry],[LineTotal],[TotalFrgn],[DebitAcctCode],[CreditAcctCode],[Currency],[OcrCode],[OcrCode2],[OcrCode3],[OcrCode4],[Dt_LastMonth],[Dscription],[Project],[U_AB_NONPROJECT]  FROM [dbo].[AB_GRPO_NON_INV] with(nolock) where [SysncSt_LastMonth]=0"
            sqlConx.Open()
            Dim sErrMsg As String = xm.ConnectSAPDB(DBName)
            If sErrMsg <> "" Then
                Functions.WriteLog(sErrMsg)
                Exit Sub
            End If
            oJE = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
            Dim nErr As Integer = 0
            Dim errMsg As String = ""
            Dim data As DataTable = GetDataSQL(query, sqlConx)
            If Not IsNothing(data) Then
             
                For Each row As DataRow In data.Rows
                    oJE.ReferenceDate = row("Dt_LastMonth")
                    oJE.Reference3 = row("DocEntry")
                    oJE.Lines.ReferenceDate1 = row("Dt_LastMonth")
                    oJE.Lines.TaxDate = row("Dt_LastMonth")
                    '[Dscription],[Project],[U_AB_NONPROJECT]
                    'If row("Project").ToString <> "" Then
                    '    oJE.Lines.ProjectCode = row("Project")
                    'End If
                    'If row("U_AB_NONPROJECT").ToString <> "" Then
                    '    oJE.Lines.UserFields.Fields.Item("U_AB_NONPROJECT").Value = row("U_AB_NONPROJECT")
                    'End If
                   
                    If row("Currency") <> "SGD" Then
                        oJE.Lines.FCCredit = row("TotalFrgn")
                        oJE.Lines.Reference1 = row("DocEntry")
                        oJE.Lines.FCCurrency = row("Currency").ToString
                        oJE.Lines.AccountCode = row("CreditAcctCode")
                        oJE.Lines.CostingCode = row("OcrCode")
                        oJE.Lines.CostingCode2 = row("OcrCode2")
                        oJE.Lines.CostingCode3 = row("OcrCode3")
                        oJE.Lines.CostingCode4 = row("OcrCode4")
                        oJE.Lines.Reference2 = row("Dscription")

                        oJE.Lines.Add()
                        'oJE.Lines.SetCurrentLine(1)

                        oJE.Lines.FCDebit = row("TotalFrgn")
                        oJE.Lines.Reference1 = row("DocEntry")
                        oJE.Lines.FCCurrency = row("Currency").ToString
                        oJE.Lines.AccountCode = row("DebitAcctCode")
                        oJE.Lines.CostingCode = row("OcrCode")
                        oJE.Lines.CostingCode2 = row("OcrCode2")
                        oJE.Lines.CostingCode3 = row("OcrCode3")
                        oJE.Lines.CostingCode4 = row("OcrCode4")
                        oJE.Lines.Reference2 = row("Dscription")

                        oJE.Lines.Add()
                    Else
                        oJE.Lines.Credit = row("LineTotal")
                        oJE.Lines.Reference1 = row("DocEntry")
                        oJE.Lines.AccountCode = row("CreditAcctCode")
                        oJE.Lines.CostingCode = row("OcrCode")
                        oJE.Lines.CostingCode2 = row("OcrCode2")
                        oJE.Lines.CostingCode3 = row("OcrCode3")
                        oJE.Lines.CostingCode4 = row("OcrCode4")
                        oJE.Lines.Reference2 = row("Dscription")

                        oJE.Lines.Add()
                        'oJE.Lines.SetCurrentLine(1)
                        oJE.Lines.Debit = row("LineTotal")
                        oJE.Lines.Reference1 = row("DocEntry")
                        oJE.Lines.AccountCode = row("DebitAcctCode")
                        oJE.Lines.CostingCode = row("OcrCode")
                        oJE.Lines.CostingCode2 = row("OcrCode2")
                        oJE.Lines.CostingCode3 = row("OcrCode3")
                        oJE.Lines.CostingCode4 = row("OcrCode4")
                        oJE.Lines.Reference2 = row("Dscription")

                        oJE.Lines.Add()
                    End If
                Next

                If (0 <> oJE.Add()) Then
                    oCompany.GetLastError(nErr, errMsg)
                    query = "update [dbo].[AB_GRPO_NON_INV] set ErrorMsg='" & errMsg.Replace("'", "''") & "' where [SysncSt_LastMonth]=0"
                    UpdateDataSQL(query, sqlConx)
                Else

                    query = "update [dbo].[AB_GRPO_NON_INV] set ErrorMsg='',[SysncSt_LastMonth]=1,[ReceiveDate_LastMonth] = getdate() where [SysncSt_LastMonth]=0"
                    UpdateDataSQL(query, sqlConx)
                End If
            End If
        Catch ex As Exception
            Functions.WriteLog(ex.Message)
        Finally
            oJE = Nothing
            sqlConx.Close()
        End Try
    End Sub
    Public Sub CreateJE_FirstMonth(ByVal ConnectionString As String, ByVal DBName As String)
        Dim oJE As SAPbobsCOM.JournalEntries
        Dim sqlConx As SqlConnection = New SqlConnection(ConnectionString)
        Try
            Dim cn As New Connection
            Dim xm As New oXML


            Dim oCompany As SAPbobsCOM.Company = PublicVariable.oCompanyInfo
            Dim query As String
            query = "SELECT [ID],[DocEntry],[LineTotal],[TotalFrgn],[DebitAcctCode],[CreditAcctCode],[Currency],[OcrCode],[OcrCode2],[OcrCode3],[OcrCode4],[Dt_FitstMonth],[Dscription],[Project],[U_AB_NONPROJECT]  FROM [dbo].[AB_GRPO_NON_INV] with(nolock) where [SysncSt_FirstMonth]=0"
            sqlConx.Open()
            Dim sErrMsg As String = xm.ConnectSAPDB(DBName)
            If sErrMsg <> "" Then
                Functions.WriteLog(sErrMsg)
                Exit Sub
            End If
            oJE = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
            Dim nErr As Integer
            Dim errMsg As String = ""

            Dim data As DataTable = GetDataSQL(query, sqlConx)
            If Not IsNothing(data) Then
                '   xm.SetDB()
             
                For Each row As DataRow In data.Rows

                    oJE.ReferenceDate = row("Dt_FitstMonth")
                    oJE.Reference3 = row("DocEntry")
                    oJE.Lines.ReferenceDate1 = row("Dt_FitstMonth")
                    oJE.Lines.TaxDate = row("Dt_FitstMonth")
                   
                    If row("Currency") <> "SGD" Then
                        oJE.Lines.FCCredit = row("TotalFrgn")
                        oJE.Lines.Reference1 = row("DocEntry")
                        oJE.Lines.FCCurrency = row("Currency").ToString
                        oJE.Lines.AccountCode = row("DebitAcctCode") 'row("CreditAcctCode")
                        oJE.Lines.CostingCode = row("OcrCode")
                        oJE.Lines.CostingCode2 = row("OcrCode2")
                        oJE.Lines.CostingCode3 = row("OcrCode3")
                        oJE.Lines.CostingCode4 = row("OcrCode4")
                        If row("Project").ToString <> "" Then
                            oJE.Lines.ProjectCode = row("Project")
                        End If
                        If row("U_AB_NONPROJECT").ToString <> "" Then
                            oJE.Lines.UserFields.Fields.Item("U_AB_NONPROJECT").Value = row("U_AB_NONPROJECT")
                        End If
                        oJE.Lines.Reference2 = row("Dscription")
                        oJE.Lines.Add()



                        oJE.Lines.FCDebit = row("TotalFrgn")
                        oJE.Lines.Reference1 = row("DocEntry")
                        oJE.Lines.FCCurrency = row("Currency").ToString
                        oJE.Lines.AccountCode = row("CreditAcctCode") 'row("DebitAcctCode")
                        oJE.Lines.CostingCode = row("OcrCode")
                        oJE.Lines.CostingCode2 = row("OcrCode2")
                        oJE.Lines.CostingCode3 = row("OcrCode3")
                        oJE.Lines.CostingCode4 = row("OcrCode4")
                        If row("Project").ToString <> "" Then
                            oJE.Lines.ProjectCode = row("Project")
                        End If
                        If row("U_AB_NONPROJECT").ToString <> "" Then
                            oJE.Lines.UserFields.Fields.Item("U_AB_NONPROJECT").Value = row("U_AB_NONPROJECT")
                        End If
                        oJE.Lines.Reference2 = row("Dscription")
                        oJE.Lines.Add()
                    Else
                        oJE.Lines.Credit = row("LineTotal")
                        oJE.Lines.Reference1 = row("DocEntry")
                        oJE.Lines.AccountCode = row("DebitAcctCode") 'row("CreditAcctCode")
                        oJE.Lines.CostingCode = row("OcrCode")
                        oJE.Lines.CostingCode2 = row("OcrCode2")
                        oJE.Lines.CostingCode3 = row("OcrCode3")
                        oJE.Lines.CostingCode4 = row("OcrCode4")
                        If row("Project").ToString <> "" Then
                            oJE.Lines.ProjectCode = row("Project")
                        End If
                        If row("U_AB_NONPROJECT").ToString <> "" Then
                            oJE.Lines.UserFields.Fields.Item("U_AB_NONPROJECT").Value = row("U_AB_NONPROJECT")
                        End If
                        oJE.Lines.Reference2 = row("Dscription")
                        oJE.Lines.Add()


                        oJE.Lines.Debit = row("LineTotal")
                        oJE.Lines.Reference1 = row("DocEntry")
                        oJE.Lines.AccountCode = row("CreditAcctCode") 'row("DebitAcctCode")
                        oJE.Lines.CostingCode = row("OcrCode")
                        oJE.Lines.CostingCode2 = row("OcrCode2")
                        oJE.Lines.CostingCode3 = row("OcrCode3")
                        oJE.Lines.CostingCode4 = row("OcrCode4")
                        If row("Project").ToString <> "" Then
                            oJE.Lines.ProjectCode = row("Project")
                        End If
                        If row("U_AB_NONPROJECT").ToString <> "" Then
                            oJE.Lines.UserFields.Fields.Item("U_AB_NONPROJECT").Value = row("U_AB_NONPROJECT")
                        End If
                        oJE.Lines.Reference2 = row("Dscription")
                        oJE.Lines.Add()

                    End If
                Next

                If (0 <> oJE.Add()) Then
                    oCompany.GetLastError(nErr, errMsg)
                    query = "update [dbo].[AB_GRPO_NON_INV] set ErrorMsg1='" & errMsg.Replace("'", "''") & "' where [SysncSt_FirstMonth]=0"
                    UpdateDataSQL(query, sqlConx)
                Else

                    query = "update [dbo].[AB_GRPO_NON_INV] set ErrorMsg1='',[SysncSt_FirstMonth]=1,[ReceiveDate_FitstMonth]= getdate() where [SysncSt_FirstMonth]=0"
                    UpdateDataSQL(query, sqlConx)
                End If
            End If
        Catch ex As Exception
            Functions.WriteLog(ex.Message)
        Finally
            oJE = Nothing
            sqlConx.Close()
        End Try
    End Sub
    Private Sub UpdateDataSQL(ByVal query As String, ByVal sqlConx As SqlConnection)
        Try
            Dim sqlCommand As SqlCommand = sqlConx.CreateCommand()
            sqlCommand.CommandText = query
            sqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Function GetDataSQL(ByVal query As String, ByVal sqlConx As SqlConnection) As DataTable
        Try
            Dim sqlAdapter As SqlDataAdapter = New SqlDataAdapter(query, sqlConx)
            Dim table As DataTable = New DataTable("GRPO")
            sqlAdapter.Fill(table)
            Return table
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Sub UpdateOutStanding(ByVal DocEntry As Integer, ByVal CardCode As String, ByVal ProjectCode As String, ByVal sqlConx As SqlConnection)
        Try
            Dim sqlCommand As SqlCommand = sqlConx.CreateCommand()
            ' Dim query As String = String.Format("Select SUM(Debit - Credit) From JDT1 where ShortName = '{0}' and Project = '{1}'", CardCode, ProjectCode)
            'sqlCommand.CommandText = query
            'Dim OutStanding As Double = Double.Parse(sqlCommand.ExecuteScalar(), Functions.GetCulture())
            'Dim OutStanding As Decimal = Decimal.Parse(sqlCommand.ExecuteScalar(), CultureInfo.InvariantCulture)
            Dim query As String = String.Format("Update OVPM Set U_Outstanding = (Select SUM(Debit - Credit) From JDT1 where ShortName = '{0}' and Project = '{1}') where DocEntry = {2}", CardCode, ProjectCode, DocEntry)
            'Functions.WriteLog(query)
            sqlCommand.CommandText = query
            sqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class

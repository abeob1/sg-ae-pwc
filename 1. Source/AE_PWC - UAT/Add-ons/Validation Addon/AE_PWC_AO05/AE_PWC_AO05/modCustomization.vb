Imports System.IO

'Imports System.IO

Module modCustomization

    Public Function GetDataViewFromTXT(ByVal CurrFileToUpload As String, ByVal Filename As String) As DataView

        ' **********************************************************************************
        '   Function    :   GetDataViewFromTXT()
        '   Purpose     :   This function will upload the data from TXT file to Dataview
        '   Parameters  :   ByRef CurrFileToUpload AS String 
        '                       CurrFileToUpload = File Name
        '   Author      :   SAI
        '   Date        :   22/1/2015
        ' **********************************************************************************

        Dim dv As DataView
        Dim iCount As Int32 = 0
        Dim sFuncName As String = String.Empty
        Dim sText As String
        Try
            sFuncName = "GetDataViewFromTXT"
            Console.WriteLine("Starting Function", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            'Console.WriteLine("Create_schema() ", sFuncName)
            'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Create_schema() ", sFuncName)
            ' Create_schema(p_oCompDef.sInboxDir, Filename)

            'The Datatable to Return
            Dim oipower As New DataTable()
            oipower.Columns.Add("AC Code", GetType(String))
            oipower.Columns.Add("Period", GetType(String))
            oipower.Columns.Add("OU Code", GetType(String))
            oipower.Columns.Add("Entity", GetType(String))
            oipower.Columns.Add("Debit And Credit Amount", GetType(String))
            oipower.Columns.Add("Amount SGD", GetType(String)) 'Amount

            'Open the file in a stream reader.
            Dim oSR As New StreamReader(CurrFileToUpload)


            Dim sString(-1) As String
            Dim sDelimiter As String() = {vbTab}

            While oSR.Peek <> -1
                sText = oSR.ReadLine()
                If Not String.IsNullOrEmpty(sText.Trim()) Then
                    sString = sText.Split(sDelimiter, StringSplitOptions.RemoveEmptyEntries) ' "RemoveEmptyEntrie" I am also using the option to remove empty entries a

                    oipower.Rows.Add(sString(0), sString(1), sString(2), sString(3), sString(4), sString(5))
                End If
            End While

            'Console.WriteLine("Del_schema() ", sFuncName)
            'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Del_schema() ", sFuncName)
            ' Del_schema(p_oCompDef.sInboxDir)

            dv = New DataView(oipower)
            Return dv

        Catch ex As Exception
            Console.WriteLine("Error occured while reading content of  " & ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error occured while reading content of  " & ex.Message, sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
            Return Nothing
        End Try

    End Function

    Public Function ImportStatistics(ByVal oForm As SAPbouiCOM.Form, ByRef sErrDesc As String, ByRef BubbleEvent As Boolean) As Long

        'Function   :   ImportStatistics()
        'Purpose    :   Import Text File Data Into UDT
        'Parameters :   ByVal oForm As SAPbouiCOM.Form
        '                   oForm=Form Type
        '               ByRef sErrDesc As String
        '                   sErrDesc=Error Description to be returned to calling function
        '               
        '                   =
        'Return     :   0 - FAILURE
        '               1 - SUCCESS
        'Author     :   SAI
        'Date       :   22/1/2015
        'Change     :

        Dim sFuncName As String = String.Empty
        Dim oDV As DataView = Nothing
        Dim oDt As DataTable
        Dim iCode As Integer
        Dim oRS As SAPbobsCOM.Recordset
        Dim sSql As String

        Try
            sFuncName = "ImportStatistics()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oRS = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling GetDataViewFromTXT Function", sFuncName)
            oDV = GetDataViewFromTXT(oForm.Items.Item("txtPath").Specific.string, p_sSelectedFileName)

            If oDV Is Nothing Then
                sErrDesc = "No Datas in the TXT file"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No Datas in the TXT file", sFuncName)
                BubbleEvent = False
                Return RTN_ERROR

            End If


            sSql = "SELECT isnull(Max(CAST( CODE as int)),0)+1 AS CODE FROM [@AB_STATITISTICSDATA]"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL :" & sSql, sFuncName)
            oRS.DoQuery(sSql)
            iCode = oRS.Fields.Item(0).Value

            oDt = oDV.Table
            sSql = String.Empty
            For Each row As DataRow In oDt.Rows
                ' write insert statement
                If row.Item(0).ToString.ToUpper.Trim() = "AB_GLCODE" Then Continue For

                sSql += " Insert Into [@AB_STATITISTICSDATA] Values ('" & iCode & "','" & iCode & "','" & row.Item(1).ToString & "','" & row.Item(2).ToString & "','" & row.Item(3).ToString & "','" & row.Item(4).ToString & "'," & CDbl(row.Item(5).ToString) & ",'" & row.Item(0).ToString & "')"

                iCode = iCode + 1
            Next

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL :" & sSql, sFuncName)

            oRS.DoQuery(sSql)

            ImportStatistics = RTN_SUCCESS

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

        Catch exc As Exception
            ImportStatistics = RTN_ERROR
            sErrDesc = exc.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        Finally
        End Try
    End Function

End Module

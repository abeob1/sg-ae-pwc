Imports System.Data
Imports System.Configuration

Namespace AE_PWC_AO06
    Module modDocDetails

        Private oEdit As SAPbouiCOM.EditText
        Private oCombo As SAPbouiCOM.ComboBox
        Private oStatic As SAPbouiCOM.StaticText
        Private oMatrix As SAPbouiCOM.Matrix
        Private objForm As SAPbouiCOM.Form
        Private oRecordSet As SAPbobsCOM.Recordset
        Private sSQL As String

        Public Sub InitializePRPOForm(ByVal sEntity As String, ByVal sDraftNo As String, ByVal sDoctype As String, ByVal sApprovedBy As String)
            Dim sFuncName As String = "Initializeform"
            Dim sErrDesc As String = String.Empty

            Try
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

                LoadFromXML("PR PO Document Display.srf", p_oSBOApplication)
                objForm = p_oSBOApplication.Forms.Item("PRPO")

                AddUserDatasources(objForm)

                oMatrix = objForm.Items.Item("47").Specific
                oMatrix.AutoResizeColumns()

                If sDoctype = "PO" Then
                    oEdit = objForm.Items.Item("12").Specific
                    oEdit.Value = sApprovedBy
                   
                ElseIf sDoctype = "PR" Then
                    oEdit = objForm.Items.Item("40").Specific
                    oEdit.Value = sApprovedBy
                    
                End If


                LoadValues(objForm, sEntity, sDraftNo, sDoctype)
                '  DisableFields(objForm)
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE

            Catch ex As Exception
                sErrDesc = ex.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                Throw New ArgumentException(sErrDesc)
            End Try
        End Sub

        Private Sub AddUserDatasources(ByVal objForm As SAPbouiCOM.Form)
            oEdit = objForm.Items.Item("4").Specific
            objForm.DataSources.UserDataSources.Add("uDocType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            oEdit.DataBind.SetBound(True, "", "uDocType")

            oEdit = objForm.Items.Item("6").Specific
            objForm.DataSources.UserDataSources.Add("uDraftNo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10)
            oEdit.DataBind.SetBound(True, "", "uDraftNo")

            oEdit = objForm.Items.Item("8").Specific
            objForm.DataSources.UserDataSources.Add("uPostDt", SAPbouiCOM.BoDataType.dt_DATE, 50)
            oEdit.DataBind.SetBound(True, "", "uPostDt")

            oEdit = objForm.Items.Item("10").Specific
            objForm.DataSources.UserDataSources.Add("uDocNo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10)
            oEdit.DataBind.SetBound(True, "", "uDocNo")

            oEdit = objForm.Items.Item("12").Specific
            objForm.DataSources.UserDataSources.Add("uPOApvr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            oEdit.DataBind.SetBound(True, "", "uPOApvr")

            oEdit = objForm.Items.Item("14").Specific
            objForm.DataSources.UserDataSources.Add("uSuppName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            oEdit.DataBind.SetBound(True, "", "uSuppName")

            oEdit = objForm.Items.Item("16").Specific
            objForm.DataSources.UserDataSources.Add("uRefNo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            oEdit.DataBind.SetBound(True, "", "uRefNo")

            oEdit = objForm.Items.Item("18").Specific
            objForm.DataSources.UserDataSources.Add("uCurr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50)
            oEdit.DataBind.SetBound(True, "", "uCurr")

            oEdit = objForm.Items.Item("20").Specific
            objForm.DataSources.UserDataSources.Add("uAmount", SAPbouiCOM.BoDataType.dt_PRICE, 50)
            'objForm.DataSources.UserDataSources.Add("uAmount", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50)
            oEdit.DataBind.SetBound(True, "", "uAmount")

            oEdit = objForm.Items.Item("22").Specific
            objForm.DataSources.UserDataSources.Add("uWaiver", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 240)
            oEdit.DataBind.SetBound(True, "", "uWaiver")

            oEdit = objForm.Items.Item("24").Specific
            objForm.DataSources.UserDataSources.Add("uBdgted", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10)
            oEdit.DataBind.SetBound(True, "", "uBdgted")

            oEdit = objForm.Items.Item("26").Specific
            objForm.DataSources.UserDataSources.Add("uBdExcd", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10)
            oEdit.DataBind.SetBound(True, "", "uBdExcd")

            oEdit = objForm.Items.Item("28").Specific
            objForm.DataSources.UserDataSources.Add("uAprlCd", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50)
            oEdit.DataBind.SetBound(True, "", "uAprlCd")

            oEdit = objForm.Items.Item("30").Specific
            objForm.DataSources.UserDataSources.Add("uPreAprd", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50)
            oEdit.DataBind.SetBound(True, "", "uPreAprd")

            oEdit = objForm.Items.Item("32").Specific
            objForm.DataSources.UserDataSources.Add("uRmks", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254)
            oEdit.DataBind.SetBound(True, "", "uRmks")

            oEdit = objForm.Items.Item("34").Specific
            objForm.DataSources.UserDataSources.Add("uEntity", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50)
            oEdit.DataBind.SetBound(True, "", "uEntity")

            oEdit = objForm.Items.Item("36").Specific
            objForm.DataSources.UserDataSources.Add("uReqstBy", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            oEdit.DataBind.SetBound(True, "", "uReqstBy")

            oEdit = objForm.Items.Item("38").Specific
            objForm.DataSources.UserDataSources.Add("uAprlUnit", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            oEdit.DataBind.SetBound(True, "", "uAprlUnit")

            oEdit = objForm.Items.Item("40").Specific
            objForm.DataSources.UserDataSources.Add("uAprdBy", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            oEdit.DataBind.SetBound(True, "", "uAprdBy")

            oEdit = objForm.Items.Item("42").Specific
            objForm.DataSources.UserDataSources.Add("uReqdDt", SAPbouiCOM.BoDataType.dt_DATE, 50)
            oEdit.DataBind.SetBound(True, "", "uReqdDt")

            oEdit = objForm.Items.Item("44").Specific
            objForm.DataSources.UserDataSources.Add("uCreator", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            oEdit.DataBind.SetBound(True, "", "uCreator")

            ''oMatrix = objForm.Items.Item("47").Specific
            ''objForm.DataSources.UserDataSources.Add("uLineId", SAPbouiCOM.BoDataType.dt_LONG_NUMBER, 10)
            ''oMatrix.Columns.Item("V_-1").DataBind.SetBound(True, "", "uLineId")

            ''objForm.DataSources.UserDataSources.Add("uItemCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30)
            ''oMatrix.Columns.Item("V_0").DataBind.SetBound(True, "", "uItemCode")

            ''objForm.DataSources.UserDataSources.Add("uItemName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            ''oMatrix.Columns.Item("V_13").DataBind.SetBound(True, "", "uItemName")

            ''objForm.DataSources.UserDataSources.Add("uGLAcct", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            ''oMatrix.Columns.Item("V_12").DataBind.SetBound(True, "", "uGLAcct")

            ''objForm.DataSources.UserDataSources.Add("uBuBdget", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            ''oMatrix.Columns.Item("V_11").DataBind.SetBound(True, "", "uBuBdget")

            ''objForm.DataSources.UserDataSources.Add("uProject", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            ''oMatrix.Columns.Item("V_10").DataBind.SetBound(True, "", "uProject")

            ''objForm.DataSources.UserDataSources.Add("uBudBal", SAPbouiCOM.BoDataType.dt_PRICE, 50)
            ''oMatrix.Columns.Item("V_9").DataBind.SetBound(True, "", "uBudBal")

            ''objForm.DataSources.UserDataSources.Add("uOper", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            ''oMatrix.Columns.Item("V_8").DataBind.SetBound(True, "", "uOper")

            ''objForm.DataSources.UserDataSources.Add("uQuantity", SAPbouiCOM.BoDataType.dt_QUANTITY, 50)
            ''oMatrix.Columns.Item("V_7").DataBind.SetBound(True, "", "uQuantity")

            ''objForm.DataSources.UserDataSources.Add("uUnitPrc", SAPbouiCOM.BoDataType.dt_PRICE, 50)
            ''oMatrix.Columns.Item("V_6").DataBind.SetBound(True, "", "uUnitPrc")

            ''objForm.DataSources.UserDataSources.Add("uLAmt", SAPbouiCOM.BoDataType.dt_PRICE, 50)
            ''oMatrix.Columns.Item("V_5").DataBind.SetBound(True, "", "uLAmt")

            ''objForm.DataSources.UserDataSources.Add("uQtAmt1", SAPbouiCOM.BoDataType.dt_PRICE, 100)
            ''oMatrix.Columns.Item("V_4").DataBind.SetBound(True, "", "uQtAmt1")

            ''objForm.DataSources.UserDataSources.Add("uQtSupNm1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            ''oMatrix.Columns.Item("V_3").DataBind.SetBound(True, "", "uQtSupNm1")

            ''objForm.DataSources.UserDataSources.Add("uQtAmt2", SAPbouiCOM.BoDataType.dt_PRICE, 100)
            ''oMatrix.Columns.Item("V_2").DataBind.SetBound(True, "", "uQtAmt2")

            ''objForm.DataSources.UserDataSources.Add("uQtSupNm2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            ''oMatrix.Columns.Item("V_1").DataBind.SetBound(True, "", "uQtSupNm2")

        End Sub

        'Private Sub LoadValues(ByVal objForm As SAPbouiCOM.Form, ByVal sEntity As String, ByVal sDraftNo As String, ByVal sDoctype As String)
        '    Dim sFuncName As String = "LoadValues"
        '    Dim sSQLUser As String = String.Empty
        '    Dim sSQLPWd As String = String.Empty
        '    Dim sTrgtDBName As String = String.Empty
        '    Dim strSQL As String = String.Empty

        '    objForm.Freeze(True)

        '    oEdit = objForm.Items.Item("4").Specific
        '    oEdit.Value = sDoctype
        '    oEdit = objForm.Items.Item("6").Specific
        '    oEdit.Value = sDraftNo
        '    oEdit = objForm.Items.Item("34").Specific
        '    oEdit.Value = sEntity
        '    oStatic = objForm.Items.Item("45").Specific
        '    oStatic.Caption = "Note : For viewing attachment please login to entity(" & sEntity & ")"


        '    sSQL = "SELECT * FROM [@AB_COMPANYDATA]  WHERE Name = '" & sEntity & "'"
        '    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL" & sSQL, sFuncName)
        '    oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        '    oRecordSet.DoQuery(sSQL)
        '    If oRecordSet.RecordCount > 0 Then
        '        sTrgtDBName = oRecordSet.Fields.Item("Name").Value

        '        If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DBUser")) Then
        '            sSQLUser = ConfigurationManager.AppSettings("DBUser")
        '        End If
        '        If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DBPwd")) Then
        '            sSQLPWd = ConfigurationManager.AppSettings("DBPwd")
        '        End If

        '        strSQL = "SELECT A.DocEntry,A.DocNum, REPLACE(CONVERT(CHAR(10), A.DocDate, 103), '/', '') DocDate,A.CardName,A.NumAtCard,A.DocCur,A.U_AB_WAIVER,A.U_AB_APPROVALAMT,A.U_AB_BudgetedCost ,A.U_AB_BudgetExceeded, " & _
        '                 " A.U_AB_APPROVALCODE ,A.U_AB_PREAPPROVED ,A.Comments,A.ReqName ,REPLACE(CONVERT(CHAR(10), A.ReqDate, 103), '/', '') ReqDate ,C.U_NAME ," & _
        '                 " B.ItemCode,B.Dscription,B.AcctCode,B.U_AB_NONPROJECT ,B.Project,B.U_AB_BALANCE,B.U_AB_OUName ,B.Quantity,B.Price,B.LineTotal, " & _
        '                 " B.U_AB_PQ1AMT,B.U_AB_PQ_SUP1,B.U_AB_PQ2AMT,B.U_AB_PQ_SUP2 " & _
        '                 " FROM " & sEntity & ".dbo.ODRF A " & _
        '                 " INNER JOIN " & sEntity & ".dbo.DRF1 B ON B.DocEntry = A.DocEntry " & _
        '                 " INNER JOIN " & sEntity & ".dbo.OUSR C ON C.USERID = A.UserSign " & _
        '                 " WHERE A.DocEntry = '" & sDraftNo & "' "

        '        oMatrix = objForm.Items.Item("47").Specific

        '        Dim oDatatable As DataTable
        '        Dim iCount As Integer = 1
        '        oDatatable = ExecuteSQLQuery_DT(strSQL, sTrgtDBName, sSQLUser, sSQLPWd)
        '        If Not oDatatable Is Nothing Then
        '            Dim oDataView As DataView = New DataView(oDatatable)

        '            For i As Integer = 0 To oDataView.Count - 1
        '                If iCount = 1 Then
        '                    oEdit = objForm.Items.Item("8").Specific
        '                    oEdit.String = oDataView(i)(2).ToString.Trim
        '                    oEdit = objForm.Items.Item("10").Specific
        '                    oEdit.Value = oDataView(i)(1).ToString.Trim
        '                    oEdit = objForm.Items.Item("14").Specific
        '                    oEdit.Value = oDataView(i)(3).ToString.Trim
        '                    oEdit = objForm.Items.Item("16").Specific
        '                    oEdit.Value = oDataView(i)(4).ToString.Trim
        '                    oEdit = objForm.Items.Item("18").Specific
        '                    oEdit.Value = oDataView(i)(5).ToString.Trim
        '                    oEdit = objForm.Items.Item("20").Specific
        '                    oEdit.Value = oDataView(i)(7).ToString.Trim
        '                    oEdit = objForm.Items.Item("22").Specific
        '                    oEdit.Value = oDataView(i)(6).ToString.Trim
        '                    oEdit = objForm.Items.Item("24").Specific
        '                    oEdit.Value = oDataView(i)(8).ToString.Trim
        '                    oEdit = objForm.Items.Item("26").Specific
        '                    oEdit.Value = oDataView(i)(9).ToString.Trim
        '                    oEdit = objForm.Items.Item("28").Specific
        '                    oEdit.Value = oDataView(i)(10).ToString.Trim
        '                    oEdit = objForm.Items.Item("30").Specific
        '                    oEdit.Value = oDataView(i)(11).ToString.Trim
        '                    oEdit = objForm.Items.Item("32").Specific
        '                    oEdit.Value = oDataView(i)(12).ToString.Trim
        '                    oEdit = objForm.Items.Item("36").Specific
        '                    oEdit.Value = oDataView(i)(13).ToString.Trim
        '                    ''  oEdit = objForm.Items.Item("42").Specific
        '                    If Not (oDataView(i)(14).ToString.Trim = String.Empty) Then
        '                        oEdit = objForm.Items.Item("42").Specific
        '                        oEdit.Value = oDataView(i)(14).ToString.Trim
        '                    End If

        '                    oEdit = objForm.Items.Item("44").Specific
        '                    oEdit.Value = oDataView(i)(15).ToString.Trim
        '                End If
        '                oMatrix.AddRow(1)
        '                oMatrix.Columns.Item("V_-1").Cells.Item(iCount).Specific.value = iCount
        '                oMatrix.Columns.Item("V_0").Cells.Item(iCount).Specific.value = oDataView(i)(16).ToString.Trim
        '                oMatrix.Columns.Item("V_13").Cells.Item(iCount).Specific.value = oDataView(i)(17).ToString.Trim
        '                oMatrix.Columns.Item("V_12").Cells.Item(iCount).Specific.value = oDataView(i)(18).ToString.Trim
        '                oMatrix.Columns.Item("V_11").Cells.Item(iCount).Specific.value = oDataView(i)(19).ToString.Trim
        '                oMatrix.Columns.Item("V_10").Cells.Item(iCount).Specific.value = oDataView(i)(20).ToString.Trim
        '                oMatrix.Columns.Item("V_9").Cells.Item(iCount).Specific.value = oDataView(i)(21).ToString.Trim
        '                oMatrix.Columns.Item("V_8").Cells.Item(iCount).Specific.value = oDataView(i)(22).ToString.Trim
        '                oMatrix.Columns.Item("V_7").Cells.Item(iCount).Specific.value = oDataView(i)(23).ToString.Trim
        '                oMatrix.Columns.Item("V_6").Cells.Item(iCount).Specific.value = oDataView(i)(24).ToString.Trim
        '                oMatrix.Columns.Item("V_5").Cells.Item(iCount).Specific.value = oDataView(i)(25).ToString.Trim
        '                oMatrix.Columns.Item("V_4").Cells.Item(iCount).Specific.value = oDataView(i)(26).ToString.Trim
        '                oMatrix.Columns.Item("V_3").Cells.Item(iCount).Specific.value = oDataView(i)(27).ToString.Trim
        '                oMatrix.Columns.Item("V_2").Cells.Item(iCount).Specific.value = oDataView(i)(28).ToString.Trim
        '                oMatrix.Columns.Item("V_1").Cells.Item(iCount).Specific.value = oDataView(i)(29).ToString.Trim
        '                iCount = iCount + 1
        '            Next
        '        End If

        '    End If

        '    objForm.Freeze(False)
        '    objForm.Update()

        'End Sub

        Private Sub LoadValues(ByVal objForm As SAPbouiCOM.Form, ByVal sEntity As String, ByVal sDraftNo As String, ByVal sDoctype As String)
            Dim sFuncName As String = "LoadValues"
            Dim sSQLUser As String = String.Empty
            Dim sSQLPWd As String = String.Empty
            Dim sTrgtDBName As String = String.Empty
            Dim strSQL As String = String.Empty

            Try
                objForm.Freeze(True)

                oEdit = objForm.Items.Item("4").Specific
                oEdit.Value = sDoctype
                oEdit = objForm.Items.Item("6").Specific
                oEdit.Value = sDraftNo
                oEdit = objForm.Items.Item("34").Specific
                oEdit.Value = sEntity
                oStatic = objForm.Items.Item("45").Specific
                oStatic.Caption = "Note : For viewing attachment please login to entity(" & sEntity & ")"

                strSQL = "SELECT A.DocEntry,A.DocNum, REPLACE(CONVERT(CHAR(10), A.DocDate, 103), '/', '') DocDate,A.CardName,A.NumAtCard,A.DocCur,A.U_AB_WAIVER,ISNULL(A.U_AB_APPROVALAMT,0.0) [U_AB_APPROVALAMT], " & _
                         " A.U_AB_BudgetedCost,A.U_AB_BudgetExceeded,A.U_AB_APPROVALCODE ,A.U_AB_PREAPPROVED ,A.Comments,A.ReqName ,REPLACE(CONVERT(CHAR(10), A.ReqDate, 103), '/', '') ReqDate ,C.U_NAME,A.U_AB_PURCHASEDEPT," & _
                         " B.ItemCode,B.Dscription,B.AcctCode,B.U_AB_NONPROJECT ,B.Project,B.U_AB_BALANCE,B.U_AB_OUName ,B.Quantity,B.Price,B.LineTotal, " & _
                         " B.U_AB_PQ1AMT,B.U_AB_PQ_SUP1,B.U_AB_PQ2AMT,B.U_AB_PQ_SUP2 " & _
                         " FROM " & sEntity & ".dbo.ODRF A " & _
                         " INNER JOIN " & sEntity & ".dbo.DRF1 B ON B.DocEntry = A.DocEntry " & _
                         " INNER JOIN " & sEntity & ".dbo.OUSR C ON C.USERID = A.UserSign " & _
                         " WHERE A.DocEntry = '" & sDraftNo & "' "

                Try
                    objForm.DataSources.DataTables.Add("DRF1")
                Catch ex As Exception
                End Try
                Dim oSApDT As SAPbouiCOM.DataTable

                oSApDT = objForm.DataSources.DataTables.Item("DRF1")
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing query " & strSQL, sFuncName)
                oSApDT.ExecuteQuery(strSQL)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Loading Line values", sFuncName)
                objForm.Items.Item("47").Specific.columns.item("V_0").databind.bind("DRF1", "ItemCode")
                objForm.Items.Item("47").Specific.columns.item("V_13").databind.bind("DRF1", "Dscription")
                objForm.Items.Item("47").Specific.columns.item("V_12").databind.bind("DRF1", "AcctCode")
                objForm.Items.Item("47").Specific.columns.item("V_11").databind.bind("DRF1", "U_AB_NONPROJECT")
                objForm.Items.Item("47").Specific.columns.item("V_10").databind.bind("DRF1", "Project")
                objForm.Items.Item("47").Specific.columns.item("V_9").databind.bind("DRF1", "U_AB_BALANCE")
                objForm.Items.Item("47").Specific.columns.item("V_8").databind.bind("DRF1", "U_AB_OUName")
                objForm.Items.Item("47").Specific.columns.item("V_7").databind.bind("DRF1", "Quantity")
                objForm.Items.Item("47").Specific.columns.item("V_6").databind.bind("DRF1", "Price")
                objForm.Items.Item("47").Specific.columns.item("V_5").databind.bind("DRF1", "LineTotal")
                objForm.Items.Item("47").Specific.columns.item("V_4").databind.bind("DRF1", "U_AB_PQ1AMT")
                objForm.Items.Item("47").Specific.columns.item("V_3").databind.bind("DRF1", "U_AB_PQ_SUP1")
                objForm.Items.Item("47").Specific.columns.item("V_2").databind.bind("DRF1", "U_AB_PQ2AMT")
                objForm.Items.Item("47").Specific.columns.item("V_1").databind.bind("DRF1", "U_AB_PQ_SUP2")
                objForm.Items.Item("47").Specific.LoadFromDataSource()
                objForm.Items.Item("47").Specific.AutoResizeColumns()

                objForm.Items.Item("8").Enabled = True
                objForm.Items.Item("42").Enabled = True

                Dim oRs As SAPbobsCOM.Recordset
                oRs = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL in Recordset " & strSQL, sFuncName)
                oRs.DoQuery(strSQL)
                If oRs.RecordCount > 0 Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Loading Header values", sFuncName)

                    oEdit = objForm.Items.Item("8").Specific
                    oEdit.String = oRs.Fields.Item("DocDate").Value
                    oEdit = objForm.Items.Item("10").Specific
                    oEdit.Value = oRs.Fields.Item("DocNum").Value
                    oEdit = objForm.Items.Item("14").Specific
                    oEdit.Value = oRs.Fields.Item("CardName").Value
                    oEdit = objForm.Items.Item("16").Specific
                    oEdit.Value = oRs.Fields.Item("NumAtCard").Value
                    oEdit = objForm.Items.Item("18").Specific
                    oEdit.Value = oRs.Fields.Item("DocCur").Value
                    oEdit = objForm.Items.Item("20").Specific
                    oEdit.Value = oRs.Fields.Item("U_AB_APPROVALAMT").Value
                    oEdit = objForm.Items.Item("22").Specific
                    oEdit.Value = oRs.Fields.Item("U_AB_WAIVER").Value
                    oEdit = objForm.Items.Item("24").Specific
                    oEdit.Value = oRs.Fields.Item("U_AB_BudgetedCost").Value
                    oEdit = objForm.Items.Item("26").Specific
                    oEdit.Value = oRs.Fields.Item("U_AB_BudgetExceeded").Value
                    oEdit = objForm.Items.Item("28").Specific
                    oEdit.Value = oRs.Fields.Item("U_AB_APPROVALCODE").Value
                    oEdit = objForm.Items.Item("30").Specific
                    oEdit.Value = oRs.Fields.Item("U_AB_PREAPPROVED").Value
                    oEdit = objForm.Items.Item("32").Specific
                    oEdit.Value = oRs.Fields.Item("Comments").Value
                    oEdit = objForm.Items.Item("36").Specific
                    oEdit.Value = oRs.Fields.Item("ReqName").Value
                    oEdit = objForm.Items.Item("38").Specific
                    oEdit.Value = oRs.Fields.Item("U_AB_PURCHASEDEPT").Value
                    oEdit = objForm.Items.Item("42").Specific
                    oEdit.String = oRs.Fields.Item("ReqDate").Value
                    oEdit = objForm.Items.Item("44").Specific
                    oEdit.Value = oRs.Fields.Item("U_NAME").Value
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs)

                objForm.Items.Item("48").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                objForm.Items.Item("8").Enabled = False
                objForm.Items.Item("42").Enabled = False

                objForm.Freeze(False)
                objForm.Update()
            Catch ex As Exception
                Call WriteToLogFile(ex.Message, sFuncName)
                objForm.Freeze(False)
                objForm.Update()
            End Try

        End Sub

        Private Sub DisableFields(ByVal objForm As SAPbouiCOM.Form)
            objForm.Items.Item("48").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            objForm.Items.Item("4").Enabled = False
            objForm.Items.Item("6").Enabled = False
            objForm.Items.Item("8").Enabled = False
            objForm.Items.Item("10").Enabled = False
            objForm.Items.Item("12").Enabled = False
            objForm.Items.Item("14").Enabled = False
            objForm.Items.Item("16").Enabled = False
            objForm.Items.Item("18").Enabled = False
            objForm.Items.Item("20").Enabled = False
            objForm.Items.Item("22").Enabled = False
            objForm.Items.Item("24").Enabled = False
            objForm.Items.Item("26").Enabled = False
            objForm.Items.Item("28").Enabled = False
            objForm.Items.Item("30").Enabled = False
            objForm.Items.Item("32").Enabled = False
            objForm.Items.Item("34").Enabled = False
            objForm.Items.Item("36").Enabled = False
            objForm.Items.Item("38").Enabled = False
            objForm.Items.Item("40").Enabled = False
            objForm.Items.Item("42").Enabled = False
            objForm.Items.Item("44").Enabled = False
            objForm.Items.Item("47").Enabled = False
        End Sub

        Public Sub DocDetails_SBO_ItemEvent(ByVal FormUID As String, ByVal pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean, ByVal objForm As SAPbouiCOM.Form)
            Dim sFuncName As String = "ApprovalWindow_SBO_ItemEvent"
            Dim sErrDesc As String = String.Empty

            Try
                If pval.Before_Action = True Then
                    Select Case pval.EventType
                        Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                            objForm = p_oSBOApplication.Forms.GetForm(pval.FormTypeEx, pval.FormTypeCount)

                    End Select
                Else
                    Select Case pval.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                            objForm = p_oSBOApplication.Forms.GetForm(pval.FormTypeEx, pval.FormTypeCount)
                            Try
                                Dim objItem, oItem As SAPbouiCOM.Item
                                oItem = objForm.Items.Item("47")
                                objItem = objForm.Items.Item("46")
                                objItem.Top = oItem.Top - 3
                                objItem.Height = oItem.Height + 5
                                objItem.Width = oItem.Width + 5

                                oMatrix = objForm.Items.Item("47").Specific
                                oMatrix.AutoResizeColumns()
                            Catch ex As Exception
                                objForm.Freeze(False)
                                objForm.Update()
                            End Try
                    End Select
                End If
            Catch ex As Exception
                sErrDesc = ex.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                Throw New ArgumentException(sErrDesc)
            End Try
        End Sub

    End Module
End Namespace


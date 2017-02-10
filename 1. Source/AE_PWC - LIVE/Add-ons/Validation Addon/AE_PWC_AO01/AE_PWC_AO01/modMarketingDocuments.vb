Option Explicit On
Imports SAPbobsCOM
Imports SAPbouiCOM
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Xml
Imports System.IO
Imports System.Windows.Forms
Imports System.Globalization
Imports System.Net.Mail


Module modMarketingDocuments

#Region "AutoBatch Processing"

    Public Function CreateAutoBatchProcess(ByVal oForm As SAPbouiCOM.Form, ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   CreateAutoBatchProcess()
        '   Purpose     :   This function will be providing to cater the processing for 
        '                   auto proceeding all items
        '                   (FIFO) Autobatch selection will be proceed the early admission date will be taking out first
        '   Parameters  :   ByVal oForm As SAPbouiCOM.Form
        '                       oForm = set the SAP UI Object Form object
        '                   ByRef sErrDesc As String
        '                       sErrDesc = Error Description to be returned to calling function
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        '   Date        :   
        '   Author      :   Sri
        '   Change      :   
        '                      
        ' **********************************************************************************

        Dim oMatrixUp As SAPbouiCOM.Matrix = Nothing
        Dim oMatrixDnLeft As SAPbouiCOM.Matrix = Nothing
        Dim oDIRecordset As SAPbobsCOM.Recordset = Nothing
        Dim sFuncName As String = String.Empty
        Dim sAdmission As String = String.Empty
        Dim sMessage As String = String.Empty
        Dim sSql As String = String.Empty
        Dim sItemCode As String = String.Empty

        Dim sDate As String = String.Empty
        Dim dDate As Date = Nothing
        Dim dMinDate As Date = Nothing

        Dim iInnerCount As Int32 = 0
        Dim iCount As Int32 = 0
        Dim lQty As Double = -1

        Try
            sFuncName = "CreateAutoBatchProcess()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Start Function", sFuncName)

            oMatrixUp = oForm.Items.Item("3").Specific
            oMatrixDnLeft = oForm.Items.Item("4").Specific

            If Not oMatrixDnLeft.Columns.Item("13").Visible Then
                sMessage = String.Format("Please be ensure to visible the [Admission Date] in {0} before proceeding ... ", oForm.Items.Item("6").Specific.Caption)
                p_oSBOApplication.MessageBox(sMessage)
                GoTo Normal_Exit
            End If

            DisplayStatus(oForm, "Please wait ..... ", "", sErrDesc)

            oForm.Freeze(True)
            For iCount = 0 To oMatrixUp.VisualRowCount - 1
                sItemCode = oMatrixUp.Columns.Item("1").Cells.Item(iCount + 1).Specific.Value
                lQty = CDbl(oMatrixUp.Columns.Item("55").Cells.Item(iCount + 1).Specific.Value)
                oMatrixUp.Columns.Item("1").Cells.Item(iCount + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                If oMatrixDnLeft.VisualRowCount > 0 And lQty <> 0.0 Then
Sort_Admission:
                    Try

                        oMatrixDnLeft.Columns.Item("13").TitleObject.Click(BoCellClickType.ct_Double)
                        oMatrixDnLeft = oForm.Items.Item("4").Specific

                        For iInnerCount = 0 To oMatrixDnLeft.VisualRowCount - 1
                            sDate = oMatrixDnLeft.Columns.Item("13").Cells.Item(iInnerCount + 1).Specific.Value
                            dDate = New Date(sDate.Substring(0, 4), sDate.Substring(4, 2), sDate.Substring(6))
                            If iInnerCount = 0 Then
                                dMinDate = dDate
                            Else
                                If dDate < dMinDate Then
                                    GoTo Sort_Admission
                                Else
                                    dMinDate = dDate
                                End If
                            End If
                        Next
                    Catch ex As Exception
                        Throw New ArgumentException(" Sorting ... " & ex.Message)
                    End Try

                    Try
                        iInnerCount = 1
                        For i = 1 To oMatrixDnLeft.VisualRowCount
                            If Convert.ToDecimal(oMatrixDnLeft.Columns.Item("3").Cells.Item(iInnerCount).Specific.Value.ToString()) - Convert.ToDecimal(oMatrixDnLeft.Columns.Item("24").Cells.Item(iInnerCount).Specific.Value.ToString()) >= lQty Then
                                oMatrixDnLeft.Columns.Item("4").Cells.Item(iInnerCount).Specific.Value = Convert.ToDecimal(lQty)
                                oForm.Items.Item("48").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                Exit For
                            Else
                                lQty = Convert.ToDecimal(Convert.ToDecimal(lQty) - (Convert.ToDecimal(oMatrixDnLeft.Columns.Item("3").Cells.Item(iInnerCount).Specific.Value.ToString()) - Convert.ToDecimal(oMatrixDnLeft.Columns.Item("24").Cells.Item(iInnerCount).Specific.Value.ToString())))
                                oMatrixDnLeft.Columns.Item("4").Cells.Item(iInnerCount).Specific.Value = (Convert.ToDecimal(oMatrixDnLeft.Columns.Item("3").Cells.Item(iInnerCount).Specific.Value.ToString()) - Convert.ToDecimal(oMatrixDnLeft.Columns.Item("24").Cells.Item(iInnerCount).Specific.Value.ToString()))
                                oForm.Items.Item("48").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            End If
                        Next


                    Catch ex As Exception
                        sErrDesc = String.Format("{0} >>  Line : {1} {2}", sItemCode, iCount, ex.Message)
                        Throw New ArgumentException(sErrDesc)
                    End Try

                End If

                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

            Next iCount
Normal_Exit:
            oForm.Freeze(False)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            CreateAutoBatchProcess = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            CreateAutoBatchProcess = RTN_ERROR
        Finally
            EndStatus(sErrDesc)
            oMatrixUp = Nothing
            oMatrixDnLeft = Nothing
            GC.Collect()  'Forces garbage collection of all generations.
        End Try
    End Function


    Public Function CreateAutoBatchProcess_PICKnPACK(ByVal oForm As SAPbouiCOM.Form, ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   CreateAutoBatchProcess_PICKnPACK()
        '   Purpose     :   This function will be providing to cater the processing for 
        '                   auto proceeding all items
        '                   (FIFO) Autobatch selection will be proceed the early admission date will be taking out first
        '   Parameters  :   ByVal oForm As SAPbouiCOM.Form
        '                       oForm = set the SAP UI Object Form object
        '                   ByRef sErrDesc As String
        '                       sErrDesc = Error Description to be returned to calling function
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        '   Date        :  
        '   Author      :   Sri
        '   Change      :   
        '                      
        ' **********************************************************************************

        Dim oMatrixUp As SAPbouiCOM.Matrix = Nothing
        Dim oMatrixDnLeft As SAPbouiCOM.Matrix = Nothing
        Dim oDIRecordset As SAPbobsCOM.Recordset = Nothing
        Dim sFuncName As String = String.Empty
        Dim sAdmission As String = String.Empty
        Dim sMessage As String = String.Empty
        Dim sSql As String = String.Empty
        Dim sItemCode As String = String.Empty

        Dim sDate As String = String.Empty
        Dim dDate As Date = Nothing
        Dim dMinDate As Date = Nothing

        Dim iInnerCount As Int32 = 0
        Dim iCount As Int32 = 0
        Dim lQty As Double = -1

        Try
            sFuncName = "CreateAutoBatchProcess()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Start Function", sFuncName)

            oMatrixUp = oForm.Items.Item("3").Specific
            oMatrixDnLeft = oForm.Items.Item("4").Specific

            If Not oMatrixDnLeft.Columns.Item("14").Visible Then
                sMessage = String.Format("Please be ensure to visible the [Manufacturing Date] in {0} before proceeding ... ", oForm.Items.Item("6").Specific.Caption)
                p_oSBOApplication.MessageBox(sMessage)
                GoTo Normal_Exit
            End If

            DisplayStatus(oForm, "Please wait ..... ", "", sErrDesc)

            oForm.Freeze(True)
            For iCount = 0 To oMatrixUp.VisualRowCount - 1
                sItemCode = oMatrixUp.Columns.Item("1").Cells.Item(iCount + 1).Specific.Value
                lQty = CDbl(oMatrixUp.Columns.Item("55").Cells.Item(iCount + 1).Specific.Value)
                oMatrixUp.Columns.Item("1").Cells.Item(iCount + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                If oMatrixDnLeft.VisualRowCount > 0 And lQty <> 0.0 Then
Sort_Admission:
                    Try

                        oMatrixDnLeft.Columns.Item("14").TitleObject.Click(BoCellClickType.ct_Double)
                        oMatrixDnLeft = oForm.Items.Item("4").Specific

                        For iInnerCount = 0 To oMatrixDnLeft.VisualRowCount - 1
                            sDate = oMatrixDnLeft.Columns.Item("14").Cells.Item(iInnerCount + 1).Specific.Value
                            dDate = New Date(sDate.Substring(0, 4), sDate.Substring(4, 2), sDate.Substring(6))
                            If iInnerCount = 0 Then
                                dMinDate = dDate
                            Else
                                If dDate < dMinDate Then
                                    GoTo Sort_Admission
                                Else
                                    dMinDate = dDate
                                End If
                            End If
                        Next
                    Catch ex As Exception
                        Throw New ArgumentException(" Sorting ... " & ex.Message)
                    End Try

                    Try

                        For iInnerCount = 1 To oMatrixDnLeft.VisualRowCount

                            'If Convert.ToDecimal(oMatrixDnLeft.Columns.Item("3").Cells.Item(1).Specific.Value.ToString()) >= lQty Then
                            '    oMatrixDnLeft.Columns.Item("4").Cells.Item(1).Specific.Value = Convert.ToDecimal(lQty)
                            '    oMatrixDnLeft.Columns.Item("1320000037").Cells.Item(1).Specific.Value = Convert.ToDecimal(lQty)
                            '    oForm.Items.Item("48").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            '    Exit For
                            'Else
                            '    lQty = Convert.ToDecimal(Convert.ToDecimal(lQty) - Convert.ToDecimal(oMatrixDnLeft.Columns.Item("3").Cells.Item(1).Specific.Value.ToString()))
                            '    oMatrixDnLeft.Columns.Item("4").Cells.Item(1).Specific.Value = oMatrixDnLeft.Columns.Item("3").Cells.Item(1).Specific.Value.ToString()
                            '    oMatrixDnLeft.Columns.Item("1320000037").Cells.Item(1).Specific.Value = oMatrixDnLeft.Columns.Item("3").Cells.Item(1).Specific.Value.ToString()
                            '    oForm.Items.Item("48").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            'End If

                            If Convert.ToDecimal(oMatrixDnLeft.Columns.Item("3").Cells.Item(iInnerCount).Specific.Value.ToString()) - Convert.ToDecimal(oMatrixDnLeft.Columns.Item("24").Cells.Item(iInnerCount).Specific.Value.ToString()) >= lQty Then
                                oMatrixDnLeft.Columns.Item("4").Cells.Item(iInnerCount).Specific.Value = Convert.ToDecimal(lQty)
                                oMatrixDnLeft.Columns.Item("1320000037").Cells.Item(iInnerCount).Specific.Value = Convert.ToDecimal(lQty)
                                oForm.Items.Item("48").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                Exit For
                            Else
                                lQty = Convert.ToDecimal(Convert.ToDecimal(lQty) - (Convert.ToDecimal(oMatrixDnLeft.Columns.Item("3").Cells.Item(iInnerCount).Specific.Value.ToString()) - Convert.ToDecimal(oMatrixDnLeft.Columns.Item("24").Cells.Item(iInnerCount).Specific.Value.ToString())))
                                oMatrixDnLeft.Columns.Item("4").Cells.Item(iInnerCount).Specific.Value = (Convert.ToDecimal(oMatrixDnLeft.Columns.Item("3").Cells.Item(iInnerCount).Specific.Value.ToString()) - Convert.ToDecimal(oMatrixDnLeft.Columns.Item("24").Cells.Item(iInnerCount).Specific.Value.ToString()))
                                oMatrixDnLeft.Columns.Item("1320000037").Cells.Item(iInnerCount).Specific.Value = (Convert.ToDecimal(oMatrixDnLeft.Columns.Item("3").Cells.Item(iInnerCount).Specific.Value.ToString()) - Convert.ToDecimal(oMatrixDnLeft.Columns.Item("24").Cells.Item(iInnerCount).Specific.Value.ToString()))
                                oForm.Items.Item("48").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            End If

                        Next
                    Catch ex As Exception
                        sErrDesc = String.Format("{0} >>  Line : {1} {2}", sItemCode, iCount, ex.Message)
                        Throw New ArgumentException(sErrDesc)
                    End Try

                End If

                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            Next iCount
Normal_Exit:
            oForm.Freeze(False)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            CreateAutoBatchProcess_PICKnPACK = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            CreateAutoBatchProcess_PICKnPACK = RTN_ERROR
        Finally
            EndStatus(sErrDesc)
            oMatrixUp = Nothing
            oMatrixDnLeft = Nothing
            GC.Collect()  'Forces garbage collection of all generations.
        End Try
    End Function


#End Region

    Public Function GetPOMaxAmount(ByRef oForm As SAPbouiCOM.Form, ByRef sErrDesc As String) As Double


        Dim sFuncName As String = String.Empty
        Dim oMatrix As SAPbouiCOM.Matrix
        Dim dMaxAmount As Double = 0
        Dim dDocAmount As Double = 0
        Dim dExchangerate As Double = 0
        Dim sDocinformation() As String
        Dim dDocDate As Date

        Try
            sFuncName = "GetPOMaxAmount"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)
            oMatrix = oForm.Items.Item("38").Specific

            For iRow As Integer = 1 To oMatrix.VisualRowCount
                If Not String.IsNullOrEmpty(oMatrix.Columns.Item("1").Cells.Item(iRow).Specific.String) Then
                    If String.IsNullOrEmpty(oMatrix.Columns.Item("14").Cells.Item(iRow).Specific.String) Then
                        oMatrix.Columns.Item("14").Cells.Item(iRow).Specific.active = True
                        sErrDesc = "Unit Price should not be Empty ...... !"
                        Return RTN_ERROR
                    End If
                End If
            Next

            sDocinformation = Split(oForm.Items.Item("22").Specific.String, " ", True) ' Document Amount Before Discount

            If sDocinformation(0) <> "SGD" Then
                dDocDate = System.DateTime.Parse(oForm.Items.Item("10").Specific.String, DateConversion, Globalization.DateTimeStyles.None)
                dExchangerate = GetExchangeRate(p_oDICompany, sDocinformation(0), dDocDate)
                dDocAmount = CDbl(sDocinformation(1)) * dExchangerate
            Else
                dDocAmount = CDbl(sDocinformation(1))
            End If
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            GetPOMaxAmount = dDocAmount
            sErrDesc = String.Empty

        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            GetPOMaxAmount = RTN_ERROR
        End Try


    End Function

    Public Function POValidation1(ByRef oForm As SAPbouiCOM.Form, ByVal dDocAmount As Double, ByRef sErrDesc As String) As Long


        Dim sFuncName As String = String.Empty
        Dim oMatrix As SAPbouiCOM.Matrix
        Dim sWavier As String = String.Empty
        Dim FsBudgeted As String = String.Empty
        Dim sSlpName As String = String.Empty

        Dim oComboSeries As SAPbouiCOM.ComboBox
        Dim sSQL As String = String.Empty
        Dim oRset As SAPbobsCOM.Recordset = Nothing

        Try
            sFuncName = "POValidation()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)
            oRset = p_oDICompany.GetBusinessObject(BoObjectTypes.BoRecordset)

            oMatrix = oForm.Items.Item("38").Specific

            Dim oNewForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount("-142", oForm.TypeCount)

            sWavier = oNewForm.Items.Item("U_AB_WAIVER").Specific.String
            If oNewForm.Items.Item("U_AB_BudgetedCost").Specific.value.ToString.ToUpper.Trim() = "Y" Then
                FsBudgeted = "Y"
            Else
                FsBudgeted = "N"
            End If

            sSQL = String.Format("SELECT T0.[SlpName] FROM OSLP T0 WHERE T0.[SlpCode]  = {0}", oForm.Items.Item("20").Specific.value.ToString.Trim())
            oRset.DoQuery(sSQL)
            sSlpName = oRset.Fields.Item("SlpName").Value

            oComboSeries = oForm.Items.Item("88").Specific

            If oComboSeries.Selected.Description.ToUpper() = "INTERNAL" Then Return RTN_SUCCESS

            'Greater then equall to 10K validation

            If dDocAmount >= 10000 Then
                If String.IsNullOrEmpty(sWavier) Then
                    For iRow As Integer = 1 To oMatrix.VisualRowCount

                        If Not String.IsNullOrEmpty(oMatrix.Columns.Item("1").Cells.Item(iRow).Specific.String) Then
                            If String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_PQ1AMT").Cells.Item(iRow).Specific.String) Or oMatrix.Columns.Item("U_AB_PQ1AMT").Cells.Item(iRow).Specific.String = 0 Then
                                p_oSBOApplication.StatusBar.SetText("Quotation 2 Amount should not be Empty ............ !" & iRow, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                Return 0
                            End If
                            If String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_PQ_SUP1").Cells.Item(iRow).Specific.String) Then
                                p_oSBOApplication.StatusBar.SetText("Quotation 2 Supplier Name should not be Empty ............ !" & iRow, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                Return 0
                            End If
                            If String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_PQ2AMT").Cells.Item(iRow).Specific.String) Or oMatrix.Columns.Item("U_AB_PQ2AMT").Cells.Item(iRow).Specific.String = 0 Then
                                p_oSBOApplication.StatusBar.SetText("Quotation 3 Amount should not be Empty ............ !" & iRow, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                Return 0
                            End If
                            If String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_PQ_SUP2").Cells.Item(iRow).Specific.String) Then
                                p_oSBOApplication.StatusBar.SetText("Quotation 3 Supplier Name should not be Empty ............ !" & iRow, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                Return 0
                            End If
                            If oMatrix.Columns.Item("U_AB_PQ1AMT").Cells.Item(iRow).Specific.String = oMatrix.Columns.Item("U_AB_PQ2AMT").Cells.Item(iRow).Specific.String Then
                                p_oSBOApplication.StatusBar.SetText("Please ensure quotes have been obtained from 3 different suppliers. Mandatory to source for 3 quotes for PO amount greater than 10k - Check line no. " & iRow, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                Return 0
                            End If
                            If oMatrix.Columns.Item("U_AB_PQ_SUP1").Cells.Item(iRow).Specific.String = oMatrix.Columns.Item("U_AB_PQ_SUP2").Cells.Item(iRow).Specific.String Then
                                p_oSBOApplication.StatusBar.SetText("Please ensure quotes have been obtained from 3 different suppliers. Mandatory to source for 3 quotes for PO amount greater than 10k - Check line no." & iRow, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                Return 0
                            End If
                        End If
                    Next

                    oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.string = dDocAmount
                Else
                    sSQL = "SELECT T0.[U_AB_APPROVALAMT] FROM " & p_sHoldingEntity & " ..[@AB_COMPETITIVEQUOTE]  T0 WHERE T0.[U_AB_From] <= " & dDocAmount & " " & _
                       "and  (T0.[U_AB_To] >= " & dDocAmount & " or T0.[U_AB_To] =0) and T0.[U_AB_BudgetedCost] = '" & FsBudgeted & "' and T0.[U_AB_PURCHDEPT] = '" & sSlpName & "'"

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting Approval Amount from [@AB_COMPETITIVEQUOTE] " & sSQL, sFuncName)
                    oRset.DoQuery(sSQL)
                    oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.string = oRset.Fields.Item("U_AB_APPROVALAMT").Value
                End If

                If Not String.IsNullOrEmpty(oForm.Items.Item("16").Specific.String) Then
                    p_oSBOApplication.StatusBar.SetText("Remarks field should not blank ........!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                    Return 0
                End If
            Else
                oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.string = dDocAmount
                '' ElseIf dDocAmount > 10000 And Not String.IsNullOrEmpty(sWavier) Then
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            POValidation1 = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message
            p_oSBOApplication.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            POValidation1 = RTN_ERROR
        Finally
            oRset = Nothing
        End Try


    End Function

    Public Function POValidation_Competitive(ByRef oForm As SAPbouiCOM.Form, ByVal dDocAmount As Double, ByRef dcomAmount As Double, ByRef sErrDesc As String) As Long


        Dim sFuncName As String = String.Empty
        Dim oMatrix As SAPbouiCOM.Matrix
        Dim sWavier As String = String.Empty
        Dim FsBudgeted As String = String.Empty
        Dim FsPreApproved As String = String.Empty
        Dim sSlpName As String = String.Empty

        Dim oComboSeries As SAPbouiCOM.ComboBox
        Dim sSQL As String = String.Empty
        Dim oRset As SAPbobsCOM.Recordset = Nothing
        Dim dQt1 As Double = 0
        Dim dqt2 As Double = 0
        Dim sqt1 As String = String.Empty
        Dim sqt2 As String = String.Empty
        Dim bCompetitive As Boolean = False

        Try
            sFuncName = "POValidation_Competitive()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)
            oRset = p_oDICompany.GetBusinessObject(BoObjectTypes.BoRecordset)

            oMatrix = oForm.Items.Item("38").Specific

            Dim oNewForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount("-142", oForm.TypeCount)
            If oNewForm.Items.Item("U_AB_BudgetedCost").Specific.value.ToString.ToUpper.Trim() = "Y" Then
                FsBudgeted = "Y"
            Else
                FsBudgeted = "N"
            End If
            sWavier = oNewForm.Items.Item("U_AB_WAIVER").Specific.String

            sSQL = String.Format("SELECT T0.[SlpName] FROM OSLP T0 WHERE T0.[SlpCode]  = {0}", oForm.Items.Item("20").Specific.value.ToString.Trim())
            oRset.DoQuery(sSQL)
            sSlpName = oRset.Fields.Item("SlpName").Value

            oComboSeries = oForm.Items.Item("88").Specific

            If oComboSeries.Selected.Description.ToUpper().Trim() = "INTERNAL" Or oComboSeries.Selected.Description.ToUpper().Trim() = "ANNUAL" Then
                oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.string = dDocAmount
                GoTo ST01
            End If

            'Greater then equall to 10K validation

            If dDocAmount >= 10000 Then
                If String.IsNullOrEmpty(sWavier) Then
                    For iRow As Integer = 1 To oMatrix.VisualRowCount

                        If Not String.IsNullOrEmpty(oMatrix.Columns.Item("1").Cells.Item(iRow).Specific.String) Then
                            If String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_PQ1AMT").Cells.Item(iRow).Specific.String) Or oMatrix.Columns.Item("U_AB_PQ1AMT").Cells.Item(iRow).Specific.String = 0 Then
                                p_oSBOApplication.StatusBar.SetText("Quotation 2 Amount should not be Empty ............ !" & iRow, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                Return 0
                            End If
                            If String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_PQ_SUP1").Cells.Item(iRow).Specific.String) Then
                                p_oSBOApplication.StatusBar.SetText("Quotation 2 Supplier Name should not be Empty ............ !" & iRow, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                Return 0
                            End If
                            If String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_PQ2AMT").Cells.Item(iRow).Specific.String) Or oMatrix.Columns.Item("U_AB_PQ2AMT").Cells.Item(iRow).Specific.String = 0 Then
                                p_oSBOApplication.StatusBar.SetText("Quotation 3 Amount should not be Empty ............ !" & iRow, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                Return 0
                            End If
                            If String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_PQ_SUP2").Cells.Item(iRow).Specific.String) Then
                                p_oSBOApplication.StatusBar.SetText("Quotation 3 Supplier Name should not be Empty ............ !" & iRow, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                Return 0
                            End If
                            If oMatrix.Columns.Item("U_AB_PQ1AMT").Cells.Item(iRow).Specific.String = oMatrix.Columns.Item("U_AB_PQ2AMT").Cells.Item(iRow).Specific.String Then
                                p_oSBOApplication.StatusBar.SetText("Please ensure quotes have been obtained from 3 different suppliers. Mandatory to source for 3 quotes for PO amount greater than 10k - Check line no. " & iRow, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                Return 0
                            End If
                            If oMatrix.Columns.Item("U_AB_PQ_SUP1").Cells.Item(iRow).Specific.String = oMatrix.Columns.Item("U_AB_PQ_SUP2").Cells.Item(iRow).Specific.String Then
                                p_oSBOApplication.StatusBar.SetText("Please ensure quotes have been obtained from 3 different suppliers. Mandatory to source for 3 quotes for PO amount greater than 10k - Check line no." & iRow, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                Return 0
                            End If
                        End If
                    Next

                    oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.string = dDocAmount
                Else
                    Dim iRow As Integer = 1

                    If Not String.IsNullOrEmpty(oMatrix.Columns.Item("1").Cells.Item(iRow).Specific.String) Then
                        If Not String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_PQ1AMT").Cells.Item(iRow).Specific.String) Then
                            dQt1 = CDbl(oMatrix.Columns.Item("U_AB_PQ1AMT").Cells.Item(iRow).Specific.String)
                        End If
                        If Not String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_PQ2AMT").Cells.Item(iRow).Specific.String) Then
                            dqt2 = CDbl(oMatrix.Columns.Item("U_AB_PQ2AMT").Cells.Item(iRow).Specific.String)
                        End If
                        If Not String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_PQ_SUP1").Cells.Item(iRow).Specific.String) Then
                            sqt1 = oMatrix.Columns.Item("U_AB_PQ_SUP1").Cells.Item(iRow).Specific.String
                        End If

                        If Not String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_PQ_SUP2").Cells.Item(iRow).Specific.String) Then
                            sqt2 = oMatrix.Columns.Item("U_AB_PQ_SUP1").Cells.Item(iRow).Specific.String
                        End If
                    End If

                    sSQL = "SELECT T0.[U_AB_APPROVALAMT] FROM " & p_sHoldingEntity & " ..[@AB_COMPETITIVEQUOTE]  T0 WHERE T0.[U_AB_From] <= " & dDocAmount & " " & _
                       "and  (T0.[U_AB_To] >= " & dDocAmount & " or T0.[U_AB_To] =0) and T0.[U_AB_BudgetedCost] = '" & FsBudgeted & "' and T0.[U_AB_PURCHDEPT] = '" & sSlpName & "'"

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting Approval Amount from [@AB_COMPETITIVEQUOTE] " & sSQL, sFuncName)
                    oRset.DoQuery(sSQL)

                    If dQt1 > 0 And dqt2 = 0 Then
                        oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.string = oRset.Fields.Item("U_AB_APPROVALAMT").Value
                    ElseIf dQt1 > 0 And dqt2 > 0 Then
                        oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.string = dDocAmount
                    ElseIf dQt1 = dqt2 Then
                        oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.string = oRset.Fields.Item("U_AB_APPROVALAMT").Value
                    Else
                        oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.string = dDocAmount
                    End If
                End If
                bCompetitive = True
ST01:

            Else
                oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.string = dDocAmount
                '' ElseIf dDocAmount > 10000 And Not String.IsNullOrEmpty(sWavier) Then
            End If

            If String.IsNullOrEmpty(oForm.Items.Item("16").Specific.String) And bCompetitive = True Then
                p_oSBOApplication.StatusBar.SetText("Remarks field should not blank ........!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                Return RTN_ERROR
            End If
            dcomAmount = CDbl(oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.string)


            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            POValidation_Competitive = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message
            p_oSBOApplication.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            POValidation_Competitive = RTN_ERROR
        Finally
            oRset = Nothing

        End Try


    End Function

    Public Function POValidation_MatrixGridCode(ByRef oForm As SAPbouiCOM.Form, ByVal dDocAmount As Double, ByRef sErrDesc As String) As Long


        Dim sFuncName As String = String.Empty
        Dim oMatrix As SAPbouiCOM.Matrix
        Dim sWavier As String = String.Empty
        Dim FsBudgeted As String = String.Empty
        Dim FsPreApproved As String = String.Empty
        Dim sSQL As String = String.Empty
        Dim oRset As SAPbobsCOM.Recordset = Nothing
        Dim sSlpName As String = String.Empty


        Try
            sFuncName = "POValidation_MatrixGridCode()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)
            oRset = p_oDICompany.GetBusinessObject(BoObjectTypes.BoRecordset)

            oMatrix = oForm.Items.Item("38").Specific

            Dim oNewForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount("-142", oForm.TypeCount)
            If oNewForm.Items.Item("U_AB_BudgetedCost").Specific.value.ToString.ToUpper.Trim() = "Y" Then
                FsBudgeted = "Y"
            Else
                FsBudgeted = "N"
            End If

            If oNewForm.Items.Item("U_AB_PREAPPROVED").Specific.value.ToString.ToUpper.Trim() = "Y" Then
                FsPreApproved = "Y"
            Else
                FsPreApproved = "N"
            End If


            sSQL = String.Format("SELECT T0.[SlpName] FROM OSLP T0 WHERE T0.[SlpCode]  = {0}", oForm.Items.Item("20").Specific.value.ToString.Trim())
            oRset.DoQuery(sSQL)
            sSlpName = oRset.Fields.Item("SlpName").Value

            If FsPreApproved = "Y" Then
                sSQL = "SELECT T0.[U_ApprGridCode] FROM " & p_sHoldingEntity & " ..[@AB_APPROVALMATRIX]  T0 WHERE T0.U_ApprDept = '" & sSlpName & "'  and T0.[U_DocType]  = 'PO' and  isnull(T0.[U_PreApprFROM],0)  <= " & dDocAmount & " and  (isnull(T0.[U_PreApprTO],0) >= " & dDocAmount & " or isnull(T0.[U_PreApprTO],0) = 0) " & _
                    "and (cast(isnull(T0.[U_PreApprFROM],0) as integer) - cast(isnull(T0.[U_PreApprTO],0) as integer)) <> 0"
            Else
                If FsBudgeted = "Y" Then
                    sSQL = "SELECT T0.[U_ApprGridCode] FROM " & p_sHoldingEntity & " ..[@AB_APPROVALMATRIX]  T0 WHERE T0.U_ApprDept = '" & sSlpName & "' and T0.[U_DocType]  = 'PO' and  isnull(T0.[U_BudFROM],0)  <= " & dDocAmount & " and  (isnull(T0.[U_BudTO],0)  >= " & dDocAmount & " or isnull(T0.[U_BudTO],0) = 0) " & _
                        "and (cast(isnull(T0.[U_BudFROM],0) as integer) - cast(isnull(T0.[U_BudTO],0) as integer)) <> 0"
                ElseIf FsBudgeted = "N" Then
                    sSQL = "SELECT T0.[U_ApprGridCode] FROM " & p_sHoldingEntity & " ..[@AB_APPROVALMATRIX]  T0 WHERE T0.U_ApprDept = '" & sSlpName & "' and T0.[U_DocType]  = 'PO' and  isnull(T0.[U_UnbudFROM],0)  <= " & dDocAmount & " and  (isnull(T0.[U_UnbudTO],0)  >= " & dDocAmount & " or isnull(T0.[U_UnbudTO],0) = 0)" & _
                        "and (cast(isnull(T0.[U_UnbudFROM],0) as integer) - cast(isnull(T0.[U_UnbudTO],0) as integer)) <> 0"
                End If
            End If
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting Approval Matrix Grid Code from [@AB_APPROVALMATRIX] " & sSQL, sFuncName)
            oRset.DoQuery(sSQL)
            If oRset.RecordCount > 0 Then
                oNewForm.Items.Item("U_AB_APPROVALCODE").Specific.string = oRset.Fields.Item("U_ApprGridCode").Value
            Else
                p_oSBOApplication.StatusBar.SetText("No valid Approval found. Please select another department. ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.String = 0
                oNewForm.Items.Item("U_AB_APPROVALCODE").Specific.String = String.Empty
                Return RTN_ERROR
            End If

            If Not String.IsNullOrEmpty(dDocAmount) Then
                If dDocAmount <= 0 Then
                    p_oSBOApplication.StatusBar.SetText("Approval amount should not be empty. Please select another department. ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.String = 0
                    oNewForm.Items.Item("U_AB_APPROVALCODE").Specific.String = ""
                    Return RTN_ERROR
                End If

            Else
                p_oSBOApplication.StatusBar.SetText("Approval amount should not be empty. Please select another department. ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.String = 0
                oNewForm.Items.Item("U_AB_APPROVALCODE").Specific.String = ""
                Return RTN_ERROR
            End If


            p_POApprovalCode = oNewForm.Items.Item("U_AB_APPROVALCODE").Specific.String

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            POValidation_MatrixGridCode = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message
            p_oSBOApplication.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            POValidation_MatrixGridCode = RTN_ERROR
        Finally
            oRset = Nothing

        End Try


    End Function

    Public Function POValidation_MatrixGridCode_Internal(ByRef oForm As SAPbouiCOM.Form, ByVal dDocAmount As Double, ByRef sErrDesc As String) As Long


        Dim sFuncName As String = String.Empty
        Dim oMatrix As SAPbouiCOM.Matrix
        Dim sWavier As String = String.Empty
        Dim FsBudgeted As String = String.Empty
        Dim FsPreApproved As String = String.Empty
        Dim sSQL As String = String.Empty
        Dim oRset As SAPbobsCOM.Recordset = Nothing
        Dim sSlpName As String = String.Empty


        Try
            sFuncName = "POValidation_MatrixGridCode_Internal()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)
            oRset = p_oDICompany.GetBusinessObject(BoObjectTypes.BoRecordset)

            oMatrix = oForm.Items.Item("38").Specific

            Dim oNewForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount("-142", oForm.TypeCount)
            If oNewForm.Items.Item("U_AB_BudgetedCost").Specific.value.ToString.ToUpper.Trim() = "Y" Then
                FsBudgeted = "Y"
            Else
                FsBudgeted = "N"
            End If


            sSQL = String.Format("SELECT T0.[SlpName] FROM OSLP T0 WHERE T0.[SlpCode]  = {0}", oForm.Items.Item("20").Specific.value.ToString.Trim())
            oRset.DoQuery(sSQL)
            sSlpName = oRset.Fields.Item("SlpName").Value

            If FsBudgeted = "Y" Then
                sSQL = "SELECT T0.[U_ApprGridCode] FROM " & p_sHoldingEntity & " ..[@AB_APPROVALMATRIX]  T0 WHERE T0.U_ApprDept = '" & sSlpName & "' and T0.[U_DocType]  = 'PO' and  isnull(T0.[U_PreApprFROM],0)  <= " & dDocAmount & " and  (isnull(T0.[U_PreApprTO],0)  >= " & dDocAmount & " or isnull(T0.[U_PreApprTO],0) = 0) " & _
                    "and (cast(isnull(T0.[U_PreApprFROM],0) as integer) - cast(isnull(T0.[U_PreApprTO],0) as integer)) <> 0"
            ElseIf FsBudgeted = "N" Then
                sSQL = "SELECT T0.[U_ApprGridCode] FROM " & p_sHoldingEntity & " ..[@AB_APPROVALMATRIX]  T0 WHERE T0.U_ApprDept = '" & sSlpName & "' and T0.[U_DocType]  = 'PO' and  isnull(T0.[U_UnbudFROM],0)  <= " & dDocAmount & " and  (isnull(T0.[U_UnbudTO],0)  >= " & dDocAmount & " or isnull(T0.[U_UnbudTO],0) = 0)" & _
                    "and (cast(isnull(T0.[U_UnbudFROM],0) as integer) - cast(isnull(T0.[U_UnbudTO],0) as integer)) <> 0"
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting Approval Matrix Grid Code from [@AB_APPROVALMATRIX] " & sSQL, sFuncName)
            oRset.DoQuery(sSQL)
            If oRset.RecordCount > 0 Then
                oNewForm.Items.Item("U_AB_APPROVALCODE").Specific.string = oRset.Fields.Item("U_ApprGridCode").Value
            Else
                p_oSBOApplication.StatusBar.SetText("No valid Approval found. Please select another department. ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.String = 0
                oNewForm.Items.Item("U_AB_APPROVALCODE").Specific.String = String.Empty
                Return RTN_ERROR
            End If

            If Not String.IsNullOrEmpty(dDocAmount) Then
                If dDocAmount <= 0 Then
                    p_oSBOApplication.StatusBar.SetText("Approval amount should not be empty. Please select another department. ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.String = 0
                    oNewForm.Items.Item("U_AB_APPROVALCODE").Specific.String = ""
                    Return RTN_ERROR
                End If

            Else
                p_oSBOApplication.StatusBar.SetText("Approval amount should not be empty. Please select another department. ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.String = 0
                oNewForm.Items.Item("U_AB_APPROVALCODE").Specific.String = ""
                Return RTN_ERROR
            End If
            p_POApprovalCode = oNewForm.Items.Item("U_AB_APPROVALCODE").Specific.String

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            POValidation_MatrixGridCode_Internal = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message
            p_oSBOApplication.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            POValidation_MatrixGridCode_Internal = RTN_ERROR
        Finally
            oRset = Nothing
        End Try

    End Function

    Public Function POValidation_Competitive_OLD(ByRef oForm As SAPbouiCOM.Form, ByVal dDocAmount As Double, ByRef dcomAmount As Double, ByRef sErrDesc As String) As Long


        Dim sFuncName As String = String.Empty
        Dim oMatrix As SAPbouiCOM.Matrix
        Dim sWavier As String = String.Empty
        Dim FsBudgeted As String = String.Empty
        Dim FsPreApproved As String = String.Empty
        Dim sSlpName As String = String.Empty

        Dim oComboSeries As SAPbouiCOM.ComboBox
        Dim sSQL As String = String.Empty
        Dim oRset As SAPbobsCOM.Recordset = Nothing
        Dim dQt1 As Double = 0
        Dim dqt2 As Double = 0
        Dim sqt1 As String = String.Empty
        Dim sqt2 As String = String.Empty
        Dim bCompetitive As Boolean = False

        Try
            sFuncName = "POValidation_Competitive()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)
            oRset = p_oDICompany.GetBusinessObject(BoObjectTypes.BoRecordset)

            oMatrix = oForm.Items.Item("38").Specific

            Dim oNewForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount("-142", oForm.TypeCount)

            sWavier = oNewForm.Items.Item("U_AB_WAIVER").Specific.String
            If oNewForm.Items.Item("U_AB_BudgetedCost").Specific.value.ToString.ToUpper.Trim() = "Y" Then
                FsBudgeted = "Y"
            Else
                FsBudgeted = "N"
            End If

            If oNewForm.Items.Item("U_AB_PREAPPROVED").Specific.value.ToString.ToUpper.Trim() = "Y" Then
                FsPreApproved = "Y"
            Else
                FsPreApproved = "N"
            End If

            sSQL = String.Format("SELECT T0.[SlpName] FROM OSLP T0 WHERE T0.[SlpCode]  = {0}", oForm.Items.Item("20").Specific.value.ToString.Trim())
            oRset.DoQuery(sSQL)
            sSlpName = oRset.Fields.Item("SlpName").Value

            oComboSeries = oForm.Items.Item("88").Specific

            If oComboSeries.Selected.Description.ToUpper() = "INTERNAL" Then
                oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.string = dDocAmount
                GoTo ST01
            End If


            'Greater then equall to 10K validation

            If dDocAmount >= 10000 Then
                If String.IsNullOrEmpty(sWavier) Then
                    For iRow As Integer = 1 To oMatrix.VisualRowCount

                        If Not String.IsNullOrEmpty(oMatrix.Columns.Item("1").Cells.Item(iRow).Specific.String) Then
                            If String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_PQ1AMT").Cells.Item(iRow).Specific.String) Or oMatrix.Columns.Item("U_AB_PQ1AMT").Cells.Item(iRow).Specific.String = 0 Then
                                p_oSBOApplication.StatusBar.SetText("Quotation 2 Amount should not be Empty ............ !" & iRow, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                Return 0
                            End If
                            If String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_PQ_SUP1").Cells.Item(iRow).Specific.String) Then
                                p_oSBOApplication.StatusBar.SetText("Quotation 2 Supplier Name should not be Empty ............ !" & iRow, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                Return 0
                            End If
                            If String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_PQ2AMT").Cells.Item(iRow).Specific.String) Or oMatrix.Columns.Item("U_AB_PQ2AMT").Cells.Item(iRow).Specific.String = 0 Then
                                p_oSBOApplication.StatusBar.SetText("Quotation 3 Amount should not be Empty ............ !" & iRow, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                Return 0
                            End If
                            If String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_PQ_SUP2").Cells.Item(iRow).Specific.String) Then
                                p_oSBOApplication.StatusBar.SetText("Quotation 3 Supplier Name should not be Empty ............ !" & iRow, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                Return 0
                            End If
                            If oMatrix.Columns.Item("U_AB_PQ1AMT").Cells.Item(iRow).Specific.String = oMatrix.Columns.Item("U_AB_PQ2AMT").Cells.Item(iRow).Specific.String Then
                                p_oSBOApplication.StatusBar.SetText("Please ensure quotes have been obtained from 3 different suppliers. Mandatory to source for 3 quotes for PO amount greater than 10k - Check line no. " & iRow, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                Return 0
                            End If
                            If oMatrix.Columns.Item("U_AB_PQ_SUP1").Cells.Item(iRow).Specific.String = oMatrix.Columns.Item("U_AB_PQ_SUP2").Cells.Item(iRow).Specific.String Then
                                p_oSBOApplication.StatusBar.SetText("Please ensure quotes have been obtained from 3 different suppliers. Mandatory to source for 3 quotes for PO amount greater than 10k - Check line no." & iRow, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                Return 0
                            End If
                        End If
                    Next

                    oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.string = dDocAmount
                Else
                    Dim iRow As Integer = 1

                    If Not String.IsNullOrEmpty(oMatrix.Columns.Item("1").Cells.Item(iRow).Specific.String) Then
                        If Not String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_PQ1AMT").Cells.Item(iRow).Specific.String) Then
                            dQt1 = CDbl(oMatrix.Columns.Item("U_AB_PQ1AMT").Cells.Item(iRow).Specific.String)
                        End If
                        If Not String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_PQ2AMT").Cells.Item(iRow).Specific.String) Then
                            dqt2 = CDbl(oMatrix.Columns.Item("U_AB_PQ2AMT").Cells.Item(iRow).Specific.String)
                        End If
                        If Not String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_PQ_SUP1").Cells.Item(iRow).Specific.String) Then
                            sqt1 = oMatrix.Columns.Item("U_AB_PQ_SUP1").Cells.Item(iRow).Specific.String
                        End If

                        If Not String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_PQ_SUP2").Cells.Item(iRow).Specific.String) Then
                            sqt2 = oMatrix.Columns.Item("U_AB_PQ_SUP1").Cells.Item(iRow).Specific.String
                        End If
                    End If

                    sSQL = "SELECT T0.[U_AB_APPROVALAMT] FROM " & p_sHoldingEntity & " ..[@AB_COMPETITIVEQUOTE]  T0 WHERE T0.[U_AB_From] <= " & dDocAmount & " " & _
                       "and  (T0.[U_AB_To] >= " & dDocAmount & " or T0.[U_AB_To] =0) and T0.[U_AB_BudgetedCost] = '" & FsBudgeted & "' and T0.[U_AB_PURCHDEPT] = '" & sSlpName & "'"

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting Approval Amount from [@AB_COMPETITIVEQUOTE] " & sSQL, sFuncName)
                    oRset.DoQuery(sSQL)

                    If dQt1 > 0 And dqt2 = 0 Then
                        oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.string = oRset.Fields.Item("U_AB_APPROVALAMT").Value
                    ElseIf dQt1 > 0 And dqt2 > 0 Then
                        oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.string = dDocAmount
                    ElseIf dQt1 = dqt2 Then
                        oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.string = oRset.Fields.Item("U_AB_APPROVALAMT").Value
                    Else
                        oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.string = dDocAmount
                    End If
                End If
                bCompetitive = True
ST01:
                dcomAmount = CDbl(oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.string)

                If FsPreApproved = "Y" Then
                    sSQL = "SELECT T0.[U_ApprGridCode] FROM " & p_sHoldingEntity & " ..[@AB_APPROVALMATRIX]  T0 WHERE T0.U_ApprDept = '" & sSlpName & "'  and T0.[U_DocType]  = 'PO' and  isnull(T0.[U_PreApprFROM],0)  <= " & dcomAmount & " and  isnull(T0.[U_PreApprTO],0) >= " & dcomAmount & ""
                Else
                    If FsBudgeted = "Y" Then
                        sSQL = "SELECT T0.[U_ApprGridCode] FROM " & p_sHoldingEntity & " ..[@AB_APPROVALMATRIX]  T0 WHERE T0.U_ApprDept = '" & sSlpName & "' and T0.[U_DocType]  = 'PO' and  isnull(T0.[U_BudFROM],0)  <= " & dcomAmount & " and  isnull(T0.[U_BudTO],0)  >= " & dcomAmount & ""
                    ElseIf FsBudgeted = "N" Then
                        sSQL = "SELECT T0.[U_ApprGridCode] FROM " & p_sHoldingEntity & " ..[@AB_APPROVALMATRIX]  T0 WHERE T0.U_ApprDept = '" & sSlpName & "' and T0.[U_DocType]  = 'PO' and  isnull(T0.[U_UnbudFROM],0)  <= " & dcomAmount & " and  isnull(T0.[U_UnbudTO],0)  >= " & dcomAmount & ""
                    End If
                End If
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting Approval Matrix Grid Code from [@AB_APPROVALMATRIX] " & sSQL, sFuncName)
                oRset.DoQuery(sSQL)
                If oRset.RecordCount > 0 Then
                    oNewForm.Items.Item("U_AB_APPROVALCODE").Specific.string = oRset.Fields.Item("U_ApprGridCode").Value
                Else
                    p_oSBOApplication.StatusBar.SetText("No valid Approval found. Please select another department. ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.String = 0
                    oNewForm.Items.Item("U_AB_APPROVALCODE").Specific.String = String.Empty
                    Return RTN_ERROR
                End If

                If Not String.IsNullOrEmpty(dcomAmount) Then
                    If dcomAmount <= 0 Then
                        p_oSBOApplication.StatusBar.SetText("Approval amount should not be empty. Please select another department. ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.String = 0
                        oNewForm.Items.Item("U_AB_APPROVALCODE").Specific.String = ""
                        Return RTN_ERROR
                    End If

                Else
                    p_oSBOApplication.StatusBar.SetText("Approval amount should not be empty. Please select another department. ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.String = 0
                    oNewForm.Items.Item("U_AB_APPROVALCODE").Specific.String = ""
                    Return RTN_ERROR
                End If

            Else
                oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.string = dDocAmount
                '' ElseIf dDocAmount > 10000 And Not String.IsNullOrEmpty(sWavier) Then
            End If

            If String.IsNullOrEmpty(oForm.Items.Item("16").Specific.String) And bCompetitive = True Then
                p_oSBOApplication.StatusBar.SetText("Remarks field should not blank ........!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                Return RTN_ERROR
            End If

            p_POApprovalCode = oNewForm.Items.Item("U_AB_APPROVALCODE").Specific.String
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            POValidation_Competitive_OLD = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message
            p_oSBOApplication.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            POValidation_Competitive_OLD = RTN_ERROR
        Finally
            oRset = Nothing

        End Try


    End Function

    Public Function Budget_Validation(ByRef oNewForm As SAPbouiCOM.Form, ByRef oForm As SAPbouiCOM.Form, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim dBAmount As Double = 0
        Dim sBAccount As String = String.Empty
        Dim sBCategory As String = String.Empty
        Dim oRow() As Data.DataRow = Nothing
        Dim dBalAmount As Double = 0.0
        Dim dComAmount As Double = 0.0
        Dim DocEntry As String = String.Empty
        Dim dCommittedAmount As Double = 0.0
        Dim dActualSpend As Double = 0.0
        Dim dBudgetAmount As Double = 0.0
        Dim oRset As SAPbobsCOM.Recordset = Nothing
        Dim sSQL As String = String.Empty
        Dim oMatrix As SAPbouiCOM.Matrix = Nothing
        Dim oDTQuantity As Data.DataTable = Nothing
        Dim oDVQuantity As DataView = Nothing
        Dim iLine As Integer = 0
        Dim iLineTotal As Double = 0
        Dim bBudgetExceeds As Boolean = False
        '' Dim bNonBudget As Boolean = False
        Dim sLineTotal As String = String.Empty
        Dim sSplit() As String

        Try
            sFuncName = "Budget_Validation()"
            oRset = p_oDICompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            oMatrix = oForm.Items.Item("38").Specific

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)
            oDVQuantity = New DataView(p_oDTPOMatrixs)
            For Each dr As DataRow In p_oDTPOMatrixs.Rows

                Select Case dr.Item("Cat").ToString.Trim()

                    Case "Prj"

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Entering into the Project", sFuncName)
                        ''  bNonBudget = True
                        dBAmount = dr.Item("LineAmount").ToString.Trim()
                        sBCategory = dr.Item("Project").ToString.Trim()
                        sBAccount = dr.Item("GLAccount").ToString.Trim()
                        iLine = dr.Item("Sno")

                        sSQL = "" & p_sHoldingEntity & " ..AE_SP005_BUDGET_COMMITTEDAMOUNT'" & p_oDICompany.CompanyDB & "','" & p_oDTConsBudget.Rows(0).Item("FinancYear") & "','True','" & sBCategory & "',''"
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling the store procedure " & sSQL, sFuncName)
                        oRset.DoQuery(sSQL)
                        dCommittedAmount = oRset.Fields.Item(0).Value
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Committed Amount  " & dCommittedAmount, sFuncName)

                        sSQL = "" & p_sHoldingEntity & " ..AE_SP006_BUDGET_ACTUALSPEND'" & p_oDICompany.CompanyDB & "','" & p_oDTConsBudget.Rows(0).Item("FinancYear") & "','True','" & sBCategory & "',''"
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling the store procedure " & sSQL, sFuncName)
                        oRset.DoQuery(sSQL)
                        dActualSpend = oRset.Fields.Item(0).Value
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Actual Amount  " & dActualSpend, sFuncName)

                        '' oRow = p_oDTConsBudget.Select("U_Account = '" & sBAccount & "' and U_PrjCode = '" & sBCategory & "'")
                        oRow = p_oDTConsBudget.Select("U_PrjCode = '" & sBCategory & "'")
                        dBudgetAmount = 0
                        If oRow.Count > 0 Then
                            For Each row As DataRow In oRow
                                dBudgetAmount += Convert.ToDouble(row("U_BudAmount"))
                            Next

                            '' dBalAmount = oRow(0)("U_BalAmount")
                            dBalAmount = dBudgetAmount - dCommittedAmount - dActualSpend
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Balance Amount  " & dBalAmount, sFuncName)
                            oDVQuantity.RowFilter = "Project='" & sBCategory & "' and Sno <" & iLine & ""
                            oDTQuantity = oDVQuantity.ToTable
                            iLineTotal = 0
                            If oDTQuantity.Rows.Count > 0 Then
                                iLineTotal = Convert.ToDecimal(oDTQuantity.Compute("sum(LineAmount)", String.Empty).ToString)
                            End If

                            oMatrix.Columns.Item("U_AB_BALANCE").Cells.Item(Convert.ToInt32(dr.Item("Sno").ToString.Trim())).Specific.String = dBalAmount - iLineTotal
                            sLineTotal = oMatrix.Columns.Item("21").Cells.Item(Convert.ToInt32(dr.Item("Sno").ToString.Trim())).Specific.String
                            sSplit = sLineTotal.Split(" ")
                            If CDbl(sSplit(1)) > (dBalAmount - iLineTotal) < 0 Then
                                bBudgetExceeds = True
                            End If
                            DocEntry = oRow(0)("DocEntry")

                            'Budget Amount from prj UDT > Line Total in the PO
                            ' With in the budget
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Budget Amount " & Str(dBalAmount) & " PO Amount " & Str(dBAmount), sFuncName)
                            If dBalAmount >= dBAmount Then
                                oRow = p_oDTPOMatrixs.Select("GLAccount = '" & sBAccount & "' and Project = '" & sBCategory & "'")
                                oRow(0)("UPdateAmount") = dBalAmount - dBAmount
                                oRow(0)("DocEntry") = DocEntry
                                'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("PO Amount With in Budget ", sFuncName)
                                'oNewForm.Items.Item("U_AB_BudgetedCost").Specific.select("N")
                            Else
                                ' Budget Exceeds
                                oRow = p_oDTPOMatrixs.Select("GLAccount = '" & sBAccount & "' and Project = '" & sBCategory & "'")
                                oRow(0)("UPdateAmount") = 0
                                oRow(0)("DocEntry") = DocEntry
                                'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("PO Amount Exceeds the Budget ", sFuncName)
                                'oNewForm.Items.Item("U_AB_BudgetedCost").Specific.select("Y")
                            End If
                        Else
                            bBudgetExceeds = True
                            oMatrix.Columns.Item("U_AB_BALANCE").Cells.Item(Convert.ToInt32(dr.Item("Sno").ToString.Trim())).Specific.String = 0
                        End If


                    Case "BU"
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Entering into the Project", sFuncName)
                        ''bNonBudget = True
                        dBAmount = dr.Item("LineAmount").ToString.Trim()
                        sBCategory = dr.Item("BU").ToString.Trim()
                        sBAccount = dr.Item("GLAccount").ToString.Trim()
                        iLine = dr.Item("Sno")

                        sSQL = "" & p_sHoldingEntity & " ..AE_SP005_BUDGET_COMMITTEDAMOUNT'" & p_oDICompany.CompanyDB & "','" & p_oDTConsBudget.Rows(0).Item("FinancYear") & "','False','" & sBCategory & "','" & sBAccount & "'"
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling the store procedure BU" & sSQL, sFuncName)
                        oRset.DoQuery(sSQL)
                        dCommittedAmount = oRset.Fields.Item(0).Value
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Committed Amount  " & dCommittedAmount, sFuncName)

                        sSQL = "" & p_sHoldingEntity & " ..AE_SP006_BUDGET_ACTUALSPEND'" & p_oDICompany.CompanyDB & "','" & p_oDTConsBudget.Rows(0).Item("FinancYear") & "','False','" & sBCategory & "','" & sBAccount & "'"
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling the store procedure BU" & sSQL, sFuncName)
                        oRset.DoQuery(sSQL)
                        dActualSpend = oRset.Fields.Item(0).Value
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Actual Amount  " & dActualSpend, sFuncName)

                        oRow = p_oDTConsBudget.Select("U_Account = '" & sBAccount & "' and U_BUCode = '" & sBCategory & "'")
                        If oRow.Count > 0 Then

                            dBudgetAmount = oRow(0)("U_BudAmount")
                            dBalAmount = dBudgetAmount - dCommittedAmount - dActualSpend
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Balance Amount  " & dBalAmount, sFuncName)
                            oDVQuantity.RowFilter = "BU ='" & sBCategory & "' and GLAccount = '" & sBAccount & "' and Sno <" & iLine & ""
                            oDTQuantity = oDVQuantity.ToTable
                            iLineTotal = 0
                            If oDTQuantity.Rows.Count > 0 Then
                                iLineTotal = Convert.ToDecimal(oDTQuantity.Compute("sum(LineAmount)", String.Empty).ToString)
                            End If
                            oMatrix.Columns.Item("U_AB_BALANCE").Cells.Item(Convert.ToInt32(dr.Item("Sno").ToString.Trim())).Specific.String = dBalAmount - iLineTotal
                            '' sLineTotal = oMatrix.Columns.Item("21").Cells.Item(Convert.ToInt32(dr.Item("Sno").ToString.Trim())).Specific.String
                            sLineTotal = oMatrix.Columns.Item("21").Cells.Item(Convert.ToInt32(dr.Item("Sno").ToString.Trim())).Specific.String
                            sSplit = sLineTotal.Split(" ")
                            If CDbl(sSplit(1)) > (dBalAmount - iLineTotal) < 0 Then
                                bBudgetExceeds = True
                            End If
                            ''dr.Item("Sno").ToString.Trim()
                            DocEntry = oRow(0)("DocEntry")

                            'consolidation Budget Amount > Line Total in the PO
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Budget Amount " & Str(dBalAmount) & " PO Amount " & Str(dBAmount), sFuncName)
                            If dBalAmount >= dBAmount Then
                                ' With in budget
                                oRow = p_oDTPOMatrixs.Select("GLAccount = '" & sBAccount & "' and BU = '" & sBCategory & "'")
                                oRow(0)("UPdateAmount") = dBalAmount - dBAmount
                                'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("PO Amount with in Budget ", sFuncName)
                                'oNewForm.Items.Item("U_AB_BudgetedCost").Specific.select("N")
                            Else
                                'Budget Exceeds
                                oRow = p_oDTPOMatrixs.Select("GLAccount = '" & sBAccount & "' and BU = '" & sBCategory & "'")
                                oRow(0)("UPdateAmount") = 0
                                oRow(0)("DocEntry") = DocEntry
                                'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Checking For U_AB_BudgetExceeded " & oNewForm.Items.Item("U_AB_BudgetExceeded").Specific.value, sFuncName)
                                'oNewForm.Items.Item("U_AB_BudgetedCost").Specific.select("Y")
                                'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("PO Amount Exceeds the Budget ", sFuncName)
                            End If
                        Else
                            oMatrix.Columns.Item("U_AB_BALANCE").Cells.Item(Convert.ToInt32(dr.Item("Sno").ToString.Trim())).Specific.String = 0
                            bBudgetExceeds = True
                        End If
                    Case Else
                        oMatrix.Columns.Item("U_AB_BALANCE").Cells.Item(Convert.ToInt32(dr.Item("Sno").ToString.Trim())).Specific.String = 0
                        bBudgetExceeds = True
                End Select
            Next


            If bBudgetExceeds = True Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Budget Flag True", sFuncName)
                oNewForm.Items.Item("U_AB_BudgetExceeded").Specific.select("Y")
                oNewForm.Items.Item("U_AB_BudgetedCost").Specific.select("N")
            Else
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Budget Flag False", sFuncName)
                oNewForm.Items.Item("U_AB_BudgetExceeded").Specific.select("N")
                oNewForm.Items.Item("U_AB_BudgetedCost").Specific.select("Y")
            End If

            Budget_Validation = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

        Catch ex As Exception
            sErrDesc = ex.Message
            p_oSBOApplication.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Budget_Validation = RTN_ERROR
        End Try
    End Function

    Public Function Budget_Validation(ByRef oForm As SAPbouiCOM.Form, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim dBAmount As Double = 0
        Dim sBAccount As String = String.Empty
        Dim sBCategory As String = String.Empty
        Dim oRow() As Data.DataRow = Nothing
        Dim dBalAmount As Double = 0.0
        Dim dComAmount As Double = 0.0
        Dim DocEntry As String = String.Empty
        Dim dCommittedAmount As Double = 0.0
        Dim dActualSpend As Double = 0.0
        Dim dBudgetAmount As Double = 0.0
        Dim oRset As SAPbobsCOM.Recordset = Nothing
        Dim sSQL As String = String.Empty
        Dim oMatrix As SAPbouiCOM.Matrix = Nothing
        Dim oDTQuantity As Data.DataTable = Nothing
        Dim oDVQuantity As DataView = Nothing
        Dim iLine As Integer = 0
        Dim iLineTotal As Double = 0
        Dim sSplit() As String
        Dim sLineTotal As String = String.Empty
        Dim bBudgetExceeds As Boolean = False

        Try
            sFuncName = "Budget_Validation()"
            oRset = p_oDICompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            oMatrix = oForm.Items.Item("38").Specific

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)
            oDVQuantity = New DataView(p_oDTPOMatrixs)
            For Each dr As DataRow In p_oDTPOMatrixs.Rows

                Select Case dr.Item("Cat").ToString.Trim()

                    Case "Prj"

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Entering into the Project", sFuncName)
                        dBAmount = dr.Item("LineAmount").ToString.Trim()
                        sBCategory = dr.Item("Project").ToString.Trim()
                        sBAccount = dr.Item("GLAccount").ToString.Trim()
                        iLine = dr.Item("Sno")

                        sSQL = "" & p_sHoldingEntity & " ..AE_SP005_BUDGET_COMMITTEDAMOUNT'" & p_sHoldingEntity & "','" & p_oDTConsBudget.Rows(0).Item("FinancYear") & "','True','" & sBCategory & "',''"
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling the store procedure " & sSQL, sFuncName)
                        oRset.DoQuery(sSQL)
                        dCommittedAmount = oRset.Fields.Item(0).Value
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Committed Amount  " & dCommittedAmount, sFuncName)

                        sSQL = "" & p_sHoldingEntity & " ..AE_SP006_BUDGET_ACTUALSPEND'" & p_sHoldingEntity & "','" & p_oDTConsBudget.Rows(0).Item("FinancYear") & "','True','" & sBCategory & "',''"
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling the store procedure " & sSQL, sFuncName)
                        oRset.DoQuery(sSQL)
                        dActualSpend = oRset.Fields.Item(0).Value
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Actual Amount  " & dActualSpend, sFuncName)

                        '' oRow = p_oDTConsBudget.Select("U_Account = '" & sBAccount & "' and U_PrjCode = '" & sBCategory & "'")
                        oRow = p_oDTConsBudget.Select("U_PrjCode = '" & sBCategory & "'")
                        dBudgetAmount = 0
                        If oRow.Count > 0 Then
                            For Each row As DataRow In oRow
                                dBudgetAmount += Convert.ToDouble(row("U_BudAmount"))
                            Next

                            '' dBalAmount = oRow(0)("U_BalAmount")
                            dBalAmount = dBudgetAmount - dCommittedAmount - dActualSpend
                            oDVQuantity.RowFilter = "Project='" & sBCategory & "' and Sno <" & iLine & ""
                            oDTQuantity = oDVQuantity.ToTable
                            iLineTotal = 0
                            If oDTQuantity.Rows.Count > 0 Then
                                iLineTotal = Convert.ToDecimal(oDTQuantity.Compute("sum(LineAmount)", String.Empty).ToString)
                            End If

                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Balance Amount  " & dBalAmount, sFuncName)
                            oMatrix.Columns.Item("U_AB_BALANCE").Cells.Item(Convert.ToInt32(dr.Item("Sno").ToString.Trim())).Specific.String = dBalAmount - iLineTotal
                            sLineTotal = oMatrix.Columns.Item("21").Cells.Item(Convert.ToInt32(dr.Item("Sno").ToString.Trim())).Specific.String
                            sSplit = sLineTotal.Split(" ")
                            If CDbl(sSplit(1)) > (dBalAmount - iLineTotal) < 0 Then
                                bBudgetExceeds = True
                            End If

                        End If


                    Case "BU"
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Entering into the Project", sFuncName)

                        dBAmount = dr.Item("LineAmount").ToString.Trim()
                        sBCategory = dr.Item("BU").ToString.Trim()
                        sBAccount = dr.Item("GLAccount").ToString.Trim()
                        iLine = dr.Item("Sno")

                        sSQL = "" & p_sHoldingEntity & " ..AE_SP005_BUDGET_COMMITTEDAMOUNT'" & p_sHoldingEntity & "','" & p_oDTConsBudget.Rows(0).Item("FinancYear") & "','False','" & sBCategory & "','" & sBAccount & "'"
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling the store procedure BU" & sSQL, sFuncName)
                        oRset.DoQuery(sSQL)
                        dCommittedAmount = oRset.Fields.Item(0).Value
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Committed Amount  " & dCommittedAmount, sFuncName)

                        sSQL = "" & p_sHoldingEntity & " ..AE_SP006_BUDGET_ACTUALSPEND'" & p_sHoldingEntity & "','" & p_oDTConsBudget.Rows(0).Item("FinancYear") & "','False','" & sBCategory & "','" & sBAccount & "'"
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling the store procedure BU" & sSQL, sFuncName)
                        oRset.DoQuery(sSQL)
                        dActualSpend = oRset.Fields.Item(0).Value
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Actual Amount  " & dActualSpend, sFuncName)

                        oRow = p_oDTConsBudget.Select("U_Account = '" & sBAccount & "' and U_BUCode = '" & sBCategory & "'")
                        If oRow.Count > 0 Then

                            dBudgetAmount = oRow(0)("U_BudAmount")
                            dBalAmount = dBudgetAmount - dCommittedAmount - dActualSpend
                            oDVQuantity.RowFilter = "BU ='" & sBCategory & "' and GLAccount = '" & sBAccount & "' and Sno <" & iLine & ""
                            oDTQuantity = oDVQuantity.ToTable
                            iLineTotal = 0
                            If oDTQuantity.Rows.Count > 0 Then
                                iLineTotal = Convert.ToDecimal(oDTQuantity.Compute("sum(LineAmount)", String.Empty).ToString)
                            End If
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Balance Amount  " & dBalAmount, sFuncName)
                            oMatrix.Columns.Item("U_AB_BALANCE").Cells.Item(Convert.ToInt32(dr.Item("Sno").ToString.Trim())).Specific.String = dBalAmount - iLineTotal
                            sLineTotal = oMatrix.Columns.Item("21").Cells.Item(Convert.ToInt32(dr.Item("Sno").ToString.Trim())).Specific.String
                            sSplit = sLineTotal.Split(" ")
                            If CDbl(sSplit(1)) > (dBalAmount - iLineTotal) < 0 Then
                                bBudgetExceeds = True
                            End If
                        End If
                End Select
            Next

            Budget_Validation = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

        Catch ex As Exception
            sErrDesc = ex.Message
            p_oSBOApplication.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Budget_Validation = RTN_ERROR
        End Try
    End Function

    Public Function Budget_Validation_OLD(ByRef oNewForm As SAPbouiCOM.Form, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim dBAmount As Double = 0
        Dim sBAccount As String = String.Empty
        Dim sBCategory As String = String.Empty
        Dim oRow() As Data.DataRow = Nothing
        Dim dBalAmount As Double = 0.0
        Dim dComAmount As Double = 0.0
        Dim DocEntry As String = String.Empty

        Try
            sFuncName = "Budget_Validation()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)
            For Each dr As DataRow In p_oDTPOMatrixs.Rows

                Select Case dr.Item("Cat").ToString.Trim()

                    Case "Prj"

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Entering into the Project", sFuncName)
                        dBAmount = dr.Item("LineAmount").ToString.Trim()
                        sBCategory = dr.Item("Project").ToString.Trim()
                        sBAccount = dr.Item("GLAccount").ToString.Trim()

                        oRow = p_oDTConsBudget.Select("U_Account = '" & sBAccount & "' and U_PrjCode = '" & sBCategory & "'")
                        If oRow.Count > 0 Then
                            dBalAmount = oRow(0)("U_BalAmount")
                            DocEntry = oRow(0)("DocEntry")

                            'Budget Amount from prj UDT > Line Total in the PO
                            ' With in the budget
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Budget Amount " & Str(dBalAmount) & " PO Amount " & Str(dBAmount), sFuncName)
                            If dBalAmount >= dBAmount Then
                                oRow = p_oDTPOMatrixs.Select("GLAccount = '" & sBAccount & "' and Project = '" & sBCategory & "'")
                                oRow(0)("UPdateAmount") = dBalAmount - dBAmount
                                oRow(0)("DocEntry") = DocEntry
                                'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Budget Exceed flag set to False ", sFuncName)
                                'oNewForm.Items.Item("U_AB_BudgetExceeded").Specific.select("N")
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("PO Amount With in Budget ", sFuncName)
                                oNewForm.Items.Item("U_AB_BudgetedCost").Specific.select("N")
                                'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Approval Amount Populated in the UDF ", sFuncName)
                                'oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.value = dAmount
                            Else
                                ' Budget Exceeds
                                oRow = p_oDTPOMatrixs.Select("GLAccount = '" & sBAccount & "' and Project = '" & sBCategory & "'")
                                oRow(0)("UPdateAmount") = 0
                                oRow(0)("DocEntry") = DocEntry
                                ' ''If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Checking For U_AB_BudgetExceeded " & oNewForm.Items.Item("U_AB_BudgetExceeded").Specific.value, sFuncName)
                                ' ''If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Budget Exceed flag set to True ", sFuncName)
                                ' ''oNewForm.Items.Item("U_AB_BudgetExceeded").Specific.select("Y")
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("PO Amount Exceeds the Budget ", sFuncName)
                                oNewForm.Items.Item("U_AB_BudgetedCost").Specific.select("Y")
                                'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Approval Amount Populated in the UDF ", sFuncName)
                                'oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.value = dAmount

                            End If
                            'Else
                            '    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Approval Amount Populated in the UDF ", sFuncName)
                            '    oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.value = dAmount
                        End If

                        ' ''Case "OU"

                        ' ''    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Entering into the Operation Unit", sFuncName)

                        ' ''    dBAmount = dr.Item("LineAmount").ToString.Trim()
                        ' ''    sBCategory = dr.Item("OU").ToString.Trim()
                        ' ''    sBAccount = dr.Item("GLAccount").ToString.Trim()

                        ' ''    oRow = p_oDTConsBudget.Select("U_Account = '" & sBAccount & "' and U_OUCode = '" & sBCategory & "'")
                        ' ''    If oRow.Count > 0 Then
                        ' ''        dBalAmount = oRow(0)("U_BalAmount")
                        ' ''        DocEntry = oRow(0)("DocEntry")

                        ' ''        'consolidation Budget Amount > Line Total in the PO
                        ' ''        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Budget Amount " & Str(dBalAmount) & " PO Amount " & Str(dBAmount), sFuncName)
                        ' ''        If dBalAmount > dBAmount Then
                        ' ''            oRow = p_oDTPOMatrixs.Select("GLAccount = '" & sBAccount & "' and OU = '" & sBCategory & "'")
                        ' ''            oRow(0)("UPdateAmount") = dBalAmount - dBAmount
                        ' ''            oRow(0)("DocEntry") = DocEntry
                        ' ''        Else
                        ' ''            oRow = p_oDTPOMatrixs.Select("GLAccount = '" & sBAccount & "' and OU = '" & sBCategory & "'")
                        ' ''            oRow(0)("UPdateAmount") = 0
                        ' ''            oRow(0)("DocEntry") = DocEntry
                        ' ''            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Checking For U_AB_BudgetExceeded " & oNewForm.Items.Item("U_AB_BudgetExceeded").Specific.value, sFuncName)
                        ' ''            If UCase(Trim(oNewForm.Items.Item("U_AB_BudgetExceeded").Specific.value)) = "N" Then
                        ' ''                oNewForm.Items.Item("U_AB_BudgetExceeded").Specific.select("Y")
                        ' ''                oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.value = dAmount
                        ' ''                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Approval Amount Populated in the UDF ", sFuncName)
                        ' ''                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Budget Exceed flag set to True ", sFuncName)
                        ' ''            End If
                        ' ''        End If
                        ' ''    End If


                    Case "BU"
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Entering into the Project", sFuncName)

                        dBAmount = dr.Item("LineAmount").ToString.Trim()
                        sBCategory = dr.Item("BU").ToString.Trim()
                        sBAccount = dr.Item("GLAccount").ToString.Trim()

                        oRow = p_oDTConsBudget.Select("U_Account = '" & sBAccount & "' and U_BUCode = '" & sBCategory & "'")
                        If oRow.Count > 0 Then
                            dBalAmount = oRow(0)("U_BalAmount")
                            DocEntry = oRow(0)("DocEntry")

                            'consolidation Budget Amount > Line Total in the PO
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Budget Amount " & Str(dBalAmount) & " PO Amount " & Str(dBAmount), sFuncName)
                            If dBalAmount >= dBAmount Then
                                ' With in budget
                                oRow = p_oDTPOMatrixs.Select("GLAccount = '" & sBAccount & "' and BU = '" & sBCategory & "'")
                                oRow(0)("UPdateAmount") = dBalAmount - dBAmount
                                'oRow(0)("DocEntry") = DocEntry
                                'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Budget Exceed flag set to False ", sFuncName)
                                'oNewForm.Items.Item("U_AB_BudgetExceeded").Specific.select("N")
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("PO Amount with in Budget ", sFuncName)
                                oNewForm.Items.Item("U_AB_BudgetedCost").Specific.select("N")
                                'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Approval Amount Populated in the UDF ", sFuncName)
                                'oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.value = dAmount
                            Else
                                'Budget Exceeds
                                oRow = p_oDTPOMatrixs.Select("GLAccount = '" & sBAccount & "' and BU = '" & sBCategory & "'")
                                oRow(0)("UPdateAmount") = 0
                                oRow(0)("DocEntry") = DocEntry
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Checking For U_AB_BudgetExceeded " & oNewForm.Items.Item("U_AB_BudgetExceeded").Specific.value, sFuncName)
                                'oNewForm.Items.Item("U_AB_BudgetExceeded").Specific.select("Y")
                                'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Budget Exceed flag set to True ", sFuncName)
                                'oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.value = dAmount
                                'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Approval Amount Populated in the UDF ", sFuncName)
                                oNewForm.Items.Item("U_AB_BudgetedCost").Specific.select("Y")
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("PO Amount Exceeds the Budget ", sFuncName)
                            End If
                            ''Else
                            ''    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Approval Amount Populated in the UDF ", sFuncName)
                            ''    oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.value = dAmount
                        End If
                End Select
            Next

            Budget_Validation_OLD = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

        Catch ex As Exception
            sErrDesc = ex.Message
            p_oSBOApplication.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Budget_Validation_OLD = RTN_ERROR
        End Try
    End Function

    Public Function POValidation_Competitive_OLD_Backup(ByRef oForm As SAPbouiCOM.Form, ByVal dDocAmount As Double, ByRef sErrDesc As String) As Long


        Dim sFuncName As String = String.Empty
        Dim oMatrix As SAPbouiCOM.Matrix
        Dim sWavier As String = String.Empty
        Dim FsBudgeted As String = String.Empty
        Dim sSlpName As String = String.Empty

        Dim oComboSeries As SAPbouiCOM.ComboBox
        Dim sSQL As String = String.Empty
        Dim oRset As SAPbobsCOM.Recordset = Nothing
        Dim dQt1 As Double = 0
        Dim dqt2 As Double = 0
        Dim sqt1 As String = String.Empty
        Dim sqt2 As String = String.Empty

        Try
            sFuncName = "POValidation_Competitive()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)
            oRset = p_oDICompany.GetBusinessObject(BoObjectTypes.BoRecordset)

            oMatrix = oForm.Items.Item("38").Specific

            Dim oNewForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount("-142", oForm.TypeCount)

            sWavier = oNewForm.Items.Item("U_AB_WAIVER").Specific.String
            If oNewForm.Items.Item("U_AB_BudgetedCost").Specific.value.ToString.ToUpper.Trim() = "Y" Then
                FsBudgeted = "Y"
            Else
                FsBudgeted = "N"
            End If

            sSQL = String.Format("SELECT T0.[SlpName] FROM OSLP T0 WHERE T0.[SlpCode]  = {0}", oForm.Items.Item("20").Specific.value.ToString.Trim())
            oRset.DoQuery(sSQL)
            sSlpName = oRset.Fields.Item("SlpName").Value

            oComboSeries = oForm.Items.Item("88").Specific

            ''  If oComboSeries.Selected.Description.ToUpper() = "INTERNAL" Then Return RTN_SUCCESS

            'Greater then equall to 10K validation

            If dDocAmount >= 10000 Then
                If String.IsNullOrEmpty(sWavier) Then
                    For iRow As Integer = 1 To oMatrix.VisualRowCount

                        If Not String.IsNullOrEmpty(oMatrix.Columns.Item("1").Cells.Item(iRow).Specific.String) Then
                            If String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_PQ1AMT").Cells.Item(iRow).Specific.String) Or oMatrix.Columns.Item("U_AB_PQ1AMT").Cells.Item(iRow).Specific.String = 0 Then
                                p_oSBOApplication.StatusBar.SetText("Quotation 2 Amount should not be Empty ............ !" & iRow, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                Return 0
                            End If
                            If String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_PQ_SUP1").Cells.Item(iRow).Specific.String) Then
                                p_oSBOApplication.StatusBar.SetText("Quotation 2 Supplier Name should not be Empty ............ !" & iRow, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                Return 0
                            End If
                            If String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_PQ2AMT").Cells.Item(iRow).Specific.String) Or oMatrix.Columns.Item("U_AB_PQ2AMT").Cells.Item(iRow).Specific.String = 0 Then
                                p_oSBOApplication.StatusBar.SetText("Quotation 3 Amount should not be Empty ............ !" & iRow, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                Return 0
                            End If
                            If String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_PQ_SUP2").Cells.Item(iRow).Specific.String) Then
                                p_oSBOApplication.StatusBar.SetText("Quotation 3 Supplier Name should not be Empty ............ !" & iRow, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                Return 0
                            End If
                            If oMatrix.Columns.Item("U_AB_PQ1AMT").Cells.Item(iRow).Specific.String = oMatrix.Columns.Item("U_AB_PQ2AMT").Cells.Item(iRow).Specific.String Then
                                p_oSBOApplication.StatusBar.SetText("Please ensure quotes have been obtained from 3 different suppliers. Mandatory to source for 3 quotes for PO amount greater than 10k - Check line no. " & iRow, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                Return 0
                            End If
                            If oMatrix.Columns.Item("U_AB_PQ_SUP1").Cells.Item(iRow).Specific.String = oMatrix.Columns.Item("U_AB_PQ_SUP2").Cells.Item(iRow).Specific.String Then
                                p_oSBOApplication.StatusBar.SetText("Please ensure quotes have been obtained from 3 different suppliers. Mandatory to source for 3 quotes for PO amount greater than 10k - Check line no." & iRow, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                Return 0
                            End If
                        End If
                    Next

                    oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.string = dDocAmount
                Else
                    Dim iRow As Integer = 1

                    If Not String.IsNullOrEmpty(oMatrix.Columns.Item("1").Cells.Item(iRow).Specific.String) Then
                        If Not String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_PQ1AMT").Cells.Item(iRow).Specific.String) Then
                            dQt1 = CDbl(oMatrix.Columns.Item("U_AB_PQ1AMT").Cells.Item(iRow).Specific.String)
                        End If
                        If Not String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_PQ2AMT").Cells.Item(iRow).Specific.String) Then
                            dqt2 = CDbl(oMatrix.Columns.Item("U_AB_PQ2AMT").Cells.Item(iRow).Specific.String)
                        End If
                        If Not String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_PQ_SUP1").Cells.Item(iRow).Specific.String) Then
                            sqt1 = oMatrix.Columns.Item("U_AB_PQ_SUP1").Cells.Item(iRow).Specific.String
                        End If

                        If Not String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_PQ_SUP2").Cells.Item(iRow).Specific.String) Then
                            sqt2 = oMatrix.Columns.Item("U_AB_PQ_SUP1").Cells.Item(iRow).Specific.String
                        End If
                    End If

                    sSQL = "SELECT T0.[U_AB_APPROVALAMT] FROM " & p_sHoldingEntity & " ..[@AB_COMPETITIVEQUOTE]  T0 WHERE T0.[U_AB_From] <= " & dDocAmount & " " & _
                       "and  (T0.[U_AB_To] >= " & dDocAmount & " or T0.[U_AB_To] =0) and T0.[U_AB_BudgetedCost] = '" & FsBudgeted & "' and T0.[U_AB_PURCHDEPT] = '" & sSlpName & "'"

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting Approval Amount from [@AB_COMPETITIVEQUOTE] " & sSQL, sFuncName)
                    oRset.DoQuery(sSQL)

                    If dQt1 > 0 And dqt2 = 0 Then
                        oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.string = oRset.Fields.Item("U_AB_APPROVALAMT").Value
                    ElseIf dQt1 > 0 And dqt2 > 0 Then
                        oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.string = dDocAmount
                    ElseIf dQt1 = dqt2 Then
                        oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.string = oRset.Fields.Item("U_AB_APPROVALAMT").Value
                    Else
                        oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.string = dDocAmount
                    End If
                End If

                If String.IsNullOrEmpty(oForm.Items.Item("16").Specific.String) Then
                    p_oSBOApplication.StatusBar.SetText("Remarks field should not blank ........!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                    Return RTN_ERROR
                End If
            Else
                oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.string = dDocAmount
                '' ElseIf dDocAmount > 10000 And Not String.IsNullOrEmpty(sWavier) Then
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            POValidation_Competitive_OLD_Backup = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message
            p_oSBOApplication.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            POValidation_Competitive_OLD_Backup = RTN_ERROR
        Finally
            oRset = Nothing
        End Try


    End Function

    Public Function POValidation(ByRef oForm As SAPbouiCOM.Form, ByVal dDocAmount As Double, ByRef sErrDesc As String) As Long


        Dim sFuncName As String = String.Empty
        Dim oMatrix As SAPbouiCOM.Matrix
        Dim sWavier As String = String.Empty
        Dim oComboSeries As SAPbouiCOM.ComboBox

        Try
            sFuncName = "POValidation()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)
            oMatrix = oForm.Items.Item("38").Specific

            Dim oNewForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount("-142", oForm.TypeCount)

            sWavier = oNewForm.Items.Item("U_AB_WAIVER").Specific.String

            oComboSeries = oForm.Items.Item("88").Specific

            If oComboSeries.Selected.Description.ToUpper() = "INTERNAL" Then Return RTN_SUCCESS

            '10000 50000
            If dDocAmount >= 10000 And dDocAmount < 50000 And String.IsNullOrEmpty(sWavier) Then
                For iRow As Integer = 1 To oMatrix.VisualRowCount

                    If Not String.IsNullOrEmpty(oMatrix.Columns.Item("1").Cells.Item(iRow).Specific.String) Then
                        If String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_PQ1AMT").Cells.Item(iRow).Specific.String) Or oMatrix.Columns.Item("U_AB_PQ1AMT").Cells.Item(iRow).Specific.String = 0 Then
                            p_oSBOApplication.StatusBar.SetText("Quotation 2 Amount should not be Empty ............ !" & iRow, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                            Return 0
                        End If
                        If String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_PQ_SUP1").Cells.Item(iRow).Specific.String) Then
                            p_oSBOApplication.StatusBar.SetText("Quotation 2 Supplier Name should not be Empty ............ !" & iRow, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                            Return 0
                        End If
                        If String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_PQ2AMT").Cells.Item(iRow).Specific.String) Or oMatrix.Columns.Item("U_AB_PQ2AMT").Cells.Item(iRow).Specific.String = 0 Then
                            p_oSBOApplication.StatusBar.SetText("Quotation 3 Amount should not be Empty ............ !" & iRow, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                            Return 0
                        End If
                        If String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_PQ_SUP2").Cells.Item(iRow).Specific.String) Then
                            p_oSBOApplication.StatusBar.SetText("Quotation 3 Supplier Name should not be Empty ............ !" & iRow, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                            Return 0
                        End If
                    End If

                Next

            ElseIf dDocAmount >= 50000 Then

                For iRow As Integer = 1 To oMatrix.VisualRowCount

                    If Not String.IsNullOrEmpty(oMatrix.Columns.Item("1").Cells.Item(iRow).Specific.String) Then
                        If String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_PQ1AMT").Cells.Item(iRow).Specific.String) Or oMatrix.Columns.Item("U_AB_PQ1AMT").Cells.Item(iRow).Specific.String = 0 Then
                            p_oSBOApplication.StatusBar.SetText("Quotation 2 Amount should not be Empty ............ !" & iRow, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                            Return 0
                        End If
                        If String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_PQ_SUP1").Cells.Item(iRow).Specific.String) Then
                            p_oSBOApplication.StatusBar.SetText("Quotation 2 Supplier Name should not be Empty ............ !" & iRow, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                            Return 0
                        End If
                        If String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_PQ2AMT").Cells.Item(iRow).Specific.String) Or oMatrix.Columns.Item("U_AB_PQ2AMT").Cells.Item(iRow).Specific.String = 0 Then
                            p_oSBOApplication.StatusBar.SetText("Quotation 3 Amount should not be Empty ............ !" & iRow, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                            Return 0
                        End If
                        If String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_PQ_SUP2").Cells.Item(iRow).Specific.String) Then
                            p_oSBOApplication.StatusBar.SetText("Quotation 3 Supplier Name should not be Empty ............ !" & iRow, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                            Return 0
                        End If
                    End If

                Next
            End If

            oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.string = dDocAmount
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            POValidation = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message
            p_oSBOApplication.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            POValidation = RTN_ERROR
        End Try


    End Function

    Public Function GetPRMaxAmount(ByRef oForm As SAPbouiCOM.Form, ByRef sErrDesc As String) As Double


        Dim sFuncName As String = String.Empty
        Dim oMatrix As SAPbouiCOM.Matrix
        Dim dMaxAmount As Double = 0
        Dim dDocAmount As Double = 0
        Dim dExchangerate As Double = 0
        Dim sDocinformation() As String
        Dim dDocDate As Date

        Try
            sFuncName = "GetPRMaxAmount"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)
            oMatrix = oForm.Items.Item("38").Specific

            For iRow As Integer = 1 To oMatrix.VisualRowCount
                If Not String.IsNullOrEmpty(oMatrix.Columns.Item("1").Cells.Item(iRow).Specific.String) Then
                    If String.IsNullOrEmpty(oMatrix.Columns.Item("14").Cells.Item(iRow).Specific.String) Then
                        oMatrix.Columns.Item("14").Cells.Item(iRow).Specific.active = True
                        sErrDesc = "Unit Price should not be Empty ...... !"
                        Return RTN_ERROR
                    End If
                End If
            Next

            sDocinformation = Split(oForm.Items.Item("22").Specific.String, " ", True) ' Document Amount Before Discount

            If sDocinformation(0) <> "SGD" Then
                dDocDate = System.DateTime.Parse(oForm.Items.Item("10").Specific.String, DateConversion, Globalization.DateTimeStyles.None)
                dExchangerate = GetExchangeRate(p_oDICompany, sDocinformation(0), dDocDate)
                dDocAmount = CDbl(sDocinformation(1)) * dExchangerate
            Else
                dDocAmount = CDbl(sDocinformation(1))
            End If
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            GetPRMaxAmount = dDocAmount
            sErrDesc = String.Empty

        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            GetPRMaxAmount = RTN_ERROR
        End Try


    End Function

End Module

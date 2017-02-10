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

    Public Function POValidation(ByRef oForm As SAPbouiCOM.Form, ByVal dDocAmount As Double, ByRef sErrDesc As String) As Long


        Dim sFuncName As String = String.Empty
        Dim oMatrix As SAPbouiCOM.Matrix
        Dim sWavier As String = String.Empty
        Dim FsBudgeted As String = String.Empty
        Dim sSlpName As String = String.Empty

        Dim dAmount1 As Double = 0
        Dim damount2 As Double = 0
        Dim fAmount1 As Boolean = False
        Dim fAmount2 As Boolean = False
        Dim fscenario1 As Boolean = False
        Dim fscenario2 As Boolean = False
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

                    For iRow As Integer = 1 To oMatrix.VisualRowCount

                        If Not String.IsNullOrEmpty(oMatrix.Columns.Item("1").Cells.Item(iRow).Specific.String) Then

                            If Not String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_PQ1AMT").Cells.Item(iRow).Specific.String) Then
                                If CDbl(oMatrix.Columns.Item("U_AB_PQ1AMT").Cells.Item(iRow).Specific.String) > 0 Then
                                    dAmount1 = CDbl(oMatrix.Columns.Item("U_AB_PQ1AMT").Cells.Item(iRow).Specific.String)
                                    fAmount1 = True
                                End If
                            End If

                            If Not String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_PQ2AMT").Cells.Item(iRow).Specific.String) Then
                                If CDbl(oMatrix.Columns.Item("U_AB_PQ2AMT").Cells.Item(iRow).Specific.String) > 0 Then
                                    damount2 = CDbl(oMatrix.Columns.Item("U_AB_PQ2AMT").Cells.Item(iRow).Specific.String)
                                    fAmount2 = True
                                End If
                            End If

                            If (fAmount1 = True And fAmount2 = False) Or (fAmount1 = False And fAmount2 = True) Then
                                p_oSBOApplication.StatusBar.SetText("Quotation 2 & Quotation 3 will contain amount otherwise make it blank ....! Line no. " & iRow, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                fAmount1 = False
                                fAmount2 = False
                                Return 0
                            End If

                            If fAmount1 = False And fAmount2 = False Then
                                fscenario1 = True
                            ElseIf fAmount1 = True And fAmount2 = True Then
                                fscenario2 = True
                            End If


                        End If
                    Next

                    If fscenario1 = True Then

                        sSQL = "SELECT T0.[U_AB_APPROVALAMT] FROM [dbo].[@AB_APPROVALMATRIXR]  T0 WHERE T0.[U_AB_From] <= " & dDocAmount & " " & _
                                             "and  (T0.[U_AB_To] >= " & dDocAmount & " or T0.[U_AB_To] =0) and T0.[U_AB_BudgetedCost] = '" & FsBudgeted & "' and T0.[U_AB_PURCHDEPT] = '" & sSlpName & "'"

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting Approval Amount from [@AB_APPROVALMATRIXR] " & sSQL, sFuncName)
                        oRset.DoQuery(sSQL)
                        oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.string = oRset.Fields.Item("U_AB_APPROVALAMT").Value
                    ElseIf fscenario2 = True Then
                        oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.string = dDocAmount
                    End If


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
            POValidation = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message
            p_oSBOApplication.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            POValidation = RTN_ERROR
        Finally
            oRset = Nothing
        End Try


    End Function

    Public Function POValidation_OLD(ByRef oForm As SAPbouiCOM.Form, ByRef sErrDesc As String) As Long


        Dim sFuncName As String = String.Empty
        Dim oMatrix As SAPbouiCOM.Matrix
        Dim sWavier As String = String.Empty
        Dim dDocAmount As Double = 0
        Dim oComboSeries As SAPbouiCOM.ComboBox

        Try
            sFuncName = "POValidation_OLD()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)
            oMatrix = oForm.Items.Item("38").Specific

            Dim oNewForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount("-142", oForm.TypeCount)
            If Not String.IsNullOrEmpty(oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.string) Then
                dDocAmount = oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.string
            Else
                dDocAmount = 0
            End If

            sWavier = oNewForm.Items.Item("U_AB_WAIVER").Specific.String

            oComboSeries = oForm.Items.Item("88").Specific

            If oComboSeries.Selected.Description.ToUpper() = "INTERNAL" Then Return RTN_SUCCESS

            '10000 50000
            If dDocAmount >= 10000 And dDocAmount < 50000 And String.IsNullOrEmpty(sWavier) Then
                For iRow As Integer = 1 To oMatrix.VisualRowCount

                    If Not String.IsNullOrEmpty(oMatrix.Columns.Item("1").Cells.Item(iRow).Specific.String) Then
                        If String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_PQ1AMT").Cells.Item(iRow).Specific.String) Or oMatrix.Columns.Item("U_AB_PQ1AMT").Cells.Item(iRow).Specific.String = 0 Then
                            p_oSBOApplication.StatusBar.SetText("Quotation 1 Amount should not be Empty ............ !" & iRow, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                            Return 0
                        End If
                        If String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_PQ_SUP1").Cells.Item(iRow).Specific.String) Then
                            p_oSBOApplication.StatusBar.SetText("Quotation 1 Supplier Name should not be Empty ............ !" & iRow, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                            Return 0
                        End If
                        If String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_PQ2AMT").Cells.Item(iRow).Specific.String) Or oMatrix.Columns.Item("U_AB_PQ2AMT").Cells.Item(iRow).Specific.String = 0 Then
                            p_oSBOApplication.StatusBar.SetText("Quotation 2 Amount should not be Empty ............ !" & iRow, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                            Return 0
                        End If
                        If String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_PQ_SUP2").Cells.Item(iRow).Specific.String) Then
                            p_oSBOApplication.StatusBar.SetText("Quotation 2 Supplier Name should not be Empty ............ !" & iRow, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                            Return 0
                        End If
                    End If

                Next

            ElseIf dDocAmount >= 50000 Then

                For iRow As Integer = 1 To oMatrix.VisualRowCount

                    If Not String.IsNullOrEmpty(oMatrix.Columns.Item("1").Cells.Item(iRow).Specific.String) Then
                        If String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_PQ1AMT").Cells.Item(iRow).Specific.String) Or oMatrix.Columns.Item("U_AB_PQ1AMT").Cells.Item(iRow).Specific.String = 0 Then
                            p_oSBOApplication.StatusBar.SetText("Quotation 1 Amount should not be Empty ............ !" & iRow, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                            Return 0
                        End If
                        If String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_PQ_SUP1").Cells.Item(iRow).Specific.String) Then
                            p_oSBOApplication.StatusBar.SetText("Quotation 1 Supplier Name should not be Empty ............ !" & iRow, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                            Return 0
                        End If
                        If String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_PQ2AMT").Cells.Item(iRow).Specific.String) Or oMatrix.Columns.Item("U_AB_PQ2AMT").Cells.Item(iRow).Specific.String = 0 Then
                            p_oSBOApplication.StatusBar.SetText("Quotation 2 Amount should not be Empty ............ !" & iRow, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                            Return 0
                        End If
                        If String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_PQ_SUP2").Cells.Item(iRow).Specific.String) Then
                            p_oSBOApplication.StatusBar.SetText("Quotation 2 Supplier Name should not be Empty ............ !" & iRow, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                            Return 0
                        End If
                    End If

                Next
            End If


            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            POValidation_OLD = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message
            p_oSBOApplication.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            POValidation_OLD = RTN_ERROR
        End Try


    End Function

End Module

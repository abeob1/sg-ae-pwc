Namespace AE_PWC_AO06
    Module modApprovalWindow

        Private oEdit As SAPbouiCOM.EditText
        Private oGrid As SAPbouiCOM.Grid
        Private oCheck As SAPbouiCOM.CheckBox
        Private oCombo As SAPbouiCOM.ComboBox
        Private oLink As SAPbouiCOM.LinkedButton
        Private oStatic As SAPbouiCOM.StaticText
        Private oRecordSet As SAPbobsCOM.Recordset
        Private bdoubleCheck As Boolean
        Private bAscOrder As Boolean
        Private objForm, oForm As SAPbouiCOM.Form
        Private sSQL, sErrorMsg As String

        Private Sub InitializeForm()
            Dim sFuncName As String = "Initializeform"
            Dim sErrDesc As String = String.Empty

            Try
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

                LoadFromXML("Approval Window.srf", p_oSBOApplication)
                objForm = p_oSBOApplication.Forms.Item("APRL")

                AddUserDatasources(objForm)
                LoadComboValues(objForm)

                'objForm.Items.Item("18").Specific.string = "t"
                objForm.Items.Item("20").Specific.string = "t"

                'oEdit = objForm.Items.Item("22").Specific
                'oEdit.Value = p_oDICompany.CompanyDB

                oForm = p_oSBOApplication.Forms.GetForm("169", 0)
                oStatic = oForm.Items.Item("8").Specific
                oEdit = objForm.Items.Item("8").Specific
                oEdit.Value = oStatic.Caption
                
                sSQL = "SELECT USER_CODE FROM OUSR WHERE U_NAME = '" & oEdit.Value & "'"
                oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecordSet.DoQuery(sSQL)
                If oRecordSet.RecordCount > 0 Then
                    oEdit = objForm.Items.Item("25").Specific
                    oEdit.Value = oRecordSet.Fields.Item("USER_CODE").Value
                Else
                    oEdit = objForm.Items.Item("25").Specific
                    oEdit.Value = String.Empty
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)

                'Choosefromlist(objForm)
                'BindChooseFromlist(objForm)

                objForm.DataSources.DataTables.Add("dtEntityList")

                LoadEmptyGrid(objForm)

                objForm.Items.Item("14").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                objForm.Items.Item("8").Enabled = False
                'objForm.Items.Item("22").Enabled = False
                objForm.Items.Item("25").Enabled = False

            Catch ex As Exception
                sErrDesc = ex.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                Throw New ArgumentException(sErrDesc)
            End Try

        End Sub

        Private Sub AddUserDatasources(ByVal objForm As SAPbouiCOM.Form)
            oCombo = objForm.Items.Item("6").Specific
            objForm.DataSources.UserDataSources.Add("uAprlType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            oCombo.DataBind.SetBound(True, "", "uAprlType")

            oEdit = objForm.Items.Item("8").Specific
            objForm.DataSources.UserDataSources.Add("uAprName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            oEdit.DataBind.SetBound(True, "", "uAprName")

            oEdit = objForm.Items.Item("10").Specific
            objForm.DataSources.UserDataSources.Add("uCreator", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            oEdit.DataBind.SetBound(True, "", "uCreator")

            oCombo = objForm.Items.Item("12").Specific
            objForm.DataSources.UserDataSources.Add("uDocs", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            oCombo.DataBind.SetBound(True, "", "uDocs")

            oEdit = objForm.Items.Item("14").Specific
            objForm.DataSources.UserDataSources.Add("uDrftNo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10)
            oEdit.DataBind.SetBound(True, "", "uDrftNo")

            oEdit = objForm.Items.Item("16").Specific
            objForm.DataSources.UserDataSources.Add("uVendor", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            oEdit.DataBind.SetBound(True, "", "uVendor")

            oEdit = objForm.Items.Item("18").Specific
            objForm.DataSources.UserDataSources.Add("uDateFrm", SAPbouiCOM.BoDataType.dt_DATE, 50)
            oEdit.DataBind.SetBound(True, "", "uDateFrm")

            oEdit = objForm.Items.Item("20").Specific
            objForm.DataSources.UserDataSources.Add("uDateTo", SAPbouiCOM.BoDataType.dt_DATE, 50)
            oEdit.DataBind.SetBound(True, "", "uDateTo")

            oEdit = objForm.Items.Item("22").Specific
            objForm.DataSources.UserDataSources.Add("uEntity", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            oEdit.DataBind.SetBound(True, "", "uEntity")

            oEdit = objForm.Items.Item("25").Specific
            objForm.DataSources.UserDataSources.Add("uAprCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            oEdit.DataBind.SetBound(True, "", "uAprCode")
        End Sub

        Private Sub LoadComboValues(ByVal objForm As SAPbouiCOM.Form)
            oCombo = objForm.Items.Item("6").Specific
            oCombo.ValidValues.Add("1", "MAIN APPROVER")
            oCombo.ValidValues.Add("2", "BACKUP APPROVER")
            oCombo.ValidValues.Add("3", "ALL")
            oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)

            oCombo = objForm.Items.Item("12").Specific
            oCombo.ValidValues.Add("ALL", "ALL")
            oCombo.ValidValues.Add("PR", "PURCHASE REQUEST")
            oCombo.ValidValues.Add("PO", "PURCHASE ORDER")
            oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        End Sub

        Private Sub LoadEmptyGrid(ByVal objForm As SAPbouiCOM.Form)
            'sSQL = "SELECT '' ENTITY,'' [DOCUMENT TYPE],'' [SELECT],'' [DRAFT NO],'' [DOCUMENT NO],'' [CREATOR NAME],'' [POSTING DATE],'' [VENDOR NAME],'' [AMOUNT(BEFORE GST)], "
            'sSQL = sSQL & " '' [APPROVAL GRID],'' APPROVALCODE,'' USERCODE,'' [STATUS],'' REMARKS,'' [APPROVED BY],'' [REASON FOR NOT APPROVING] "
            sSQL = "SELECT '' [APPROVAL TYPE],'' [ENTITY NAME],'' [DOCUMENT TYPE],'' [SELECT],'' [DRAFT NO],'' [DOCUMENT NO],'' [POSTING DATE],'' [VENDOR NAME],'' [AMOUNT(BEFORE GST)], "
            sSQL = sSQL & " ''[APPROVED BY],'' [APPROVAL GRID],'' [STATUS],'' REMARKS,'' [CREATOR NAME],'' [REASON FOR NOT APPROVING] "
            oGrid = objForm.Items.Item("23").Specific
            objForm.DataSources.DataTables.Item("dtEntityList").Rows.Clear()
            objForm.DataSources.DataTables.Item("dtEntityList").ExecuteQuery(sSQL)
            oGrid.DataTable = objForm.DataSources.DataTables.Item("dtEntityList")
        End Sub

        Private Sub LoadGrid(ByVal objForm As SAPbouiCOM.Form)
            Dim oColumn As SAPbouiCOM.EditTextColumn
            Dim sApproverType, sApproverCode, sApprover, sCreator, sDocument, sDraftNo, sVendor, sEntity As String
            Dim dtFromDate, dtToDate As Date
            Dim oRecordSet As SAPbobsCOM.Recordset
            Dim sFuncName As String = String.Empty
            Dim sErrDesc As String = String.Empty

            Try
                sFuncName = "LoadGrid()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

                oEdit = objForm.Items.Item("8").Specific
                sApprover = oEdit.Value

                sSQL = "SELECT USER_CODE FROM OUSR WHERE U_NAME = '" & sApprover & "'"
                oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecordSet.DoQuery(sSQL)
                If oRecordSet.RecordCount > 0 Then
                    sApproverCode = oRecordSet.Fields.Item("USER_CODE").Value
                Else
                    sApproverCode = String.Empty
                End If

                oCombo = objForm.Items.Item("6").Specific
                If oCombo.Selected.Value = "3" Then
                    sApproverType = "ALL"
                ElseIf oCombo.Selected.Value = "1" Then
                    sApproverType = "MAIN"
                ElseIf oCombo.Selected.Value = "2" Then
                    sApproverType = "BACKUP"
                End If

                oEdit = objForm.Items.Item("10").Specific
                sCreator = oEdit.Value
                oCombo = objForm.Items.Item("12").Specific
                sDocument = oCombo.Selected.Value
                oEdit = objForm.Items.Item("14").Specific
                sDraftNo = oEdit.Value
                oEdit = objForm.Items.Item("16").Specific
                sVendor = oEdit.Value
                oEdit = objForm.Items.Item("22").Specific
                sEntity = oEdit.Value
                oEdit = objForm.Items.Item("18").Specific
                dtFromDate = GetDateTimeValue(oEdit.String)
                oEdit = objForm.Items.Item("20").Specific
                dtToDate = GetDateTimeValue(oEdit.String)

                sSQL = "EXEC AE_APPROVALGRID '" & sHoldingDB & "','" & sApproverType & "','" & sApproverCode & "','" & sApprover & "','" & sCreator & "','" & sDocument & "','" & sDraftNo & "','" & sVendor & "','" & dtFromDate.ToString("yyyy-MM-dd") & "','" & dtToDate.ToString("yyyy-MM-dd") & "','" & sEntity & "'"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Extracting the Pending Documents " & sSQL, sFuncName)
                oGrid = objForm.Items.Item("23").Specific
                objForm.DataSources.DataTables.Item("dtEntityList").Rows.Clear()
                objForm.DataSources.DataTables.Item("dtEntityList").ExecuteQuery(sSQL)
                oGrid.DataTable = objForm.DataSources.DataTables.Item("dtEntityList")

                'oGrid.Columns.Item("DOCUMENT NO").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                oColumn = oGrid.Columns.Item("DRAFT NO")
                oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_Drafts

                oGrid.Columns.Item("SELECT").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
                oGrid.Columns.Item("ENTITY").Editable = False
                oGrid.Columns.Item("ENTITY NAME").Editable = False
                oGrid.Columns.Item("DOCUMENT TYPE").Editable = False
                oGrid.Columns.Item("DRAFT NO").Editable = False
                oGrid.Columns.Item("DOCUMENT NO").Editable = False
                oGrid.Columns.Item("CREATOR NAME").Editable = False
                oGrid.Columns.Item("POSTING DATE").Editable = False
                oGrid.Columns.Item("VENDOR NAME").Editable = False
                oGrid.Columns.Item("AMOUNT(BEFORE GST)").Editable = False
                oGrid.Columns.Item("APPROVAL GRID").Editable = False
                oGrid.Columns.Item("STATUS").Editable = False
                oGrid.Columns.Item("REMARKS").Editable = False
                oGrid.Columns.Item("APPROVED BY").Editable = False
                oGrid.Columns.Item("APPROVAL TYPE").Editable = False

                oGrid.Columns.Item("ENTITY").Visible = False
                oGrid.Columns.Item("APPROVALCODE").Visible = False
                oGrid.Columns.Item("USERCODE").Visible = False

                oGrid.Columns.Item("ENTITY").TitleObject.Sortable = True
                oGrid.Columns.Item("ENTITY NAME").TitleObject.Sortable = True
                oGrid.Columns.Item("DOCUMENT TYPE").TitleObject.Sortable = True
                oGrid.Columns.Item("DRAFT NO").TitleObject.Sortable = True
                oGrid.Columns.Item("DOCUMENT NO").TitleObject.Sortable = True
                oGrid.Columns.Item("CREATOR NAME").TitleObject.Sortable = True
                oGrid.Columns.Item("POSTING DATE").TitleObject.Sortable = True
                oGrid.Columns.Item("VENDOR NAME").TitleObject.Sortable = True
                oGrid.Columns.Item("AMOUNT(BEFORE GST)").TitleObject.Sortable = True
                oGrid.Columns.Item("APPROVAL GRID").TitleObject.Sortable = True
                oGrid.Columns.Item("STATUS").TitleObject.Sortable = True
                oGrid.Columns.Item("REMARKS").TitleObject.Sortable = True
                oGrid.Columns.Item("APPROVED BY").TitleObject.Sortable = True

            Catch ex As Exception
                sErrDesc = ex.Message
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                WriteToLogFile(Err.Description, sFuncName)
                ShowErr(sErrDesc)
            End Try
           
        End Sub

        Private Sub Choosefromlist(ByVal objForm As SAPbouiCOM.Form)
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            oCFLs = objForm.ChooseFromLists
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            oCFLCreationParams = p_oSBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.MultiSelection = False

            'Creator Name
            oCFLCreationParams.UniqueID = "CFL1"
            oCFLCreationParams.ObjectType = "12"
            oCFL = oCFLs.Add(oCFLCreationParams)

            'Vendor Name
            oCFLCreationParams.UniqueID = "CFL2"
            oCFLCreationParams.ObjectType = "2"
            oCFL = oCFLs.Add(oCFLCreationParams)

        End Sub

        Private Sub BindChooseFromlist(ByVal objForm As SAPbouiCOM.Form)
            oEdit = objForm.Items.Item("10").Specific
            oEdit.ChooseFromListUID = "CFL1"
            oEdit.ChooseFromListAlias = "U_Name"

            oEdit = objForm.Items.Item("16").Specific
            oEdit.ChooseFromListUID = "CFL2"
            oEdit.ChooseFromListAlias = "CardName"

        End Sub

        Private Function CheckFields(ByVal objForm As SAPbouiCOM.Form) As Boolean
            Dim sSelect As String
            Dim v_check As Boolean
            v_check = True

            oGrid = objForm.Items.Item("23").Specific
            For i = 0 To oGrid.DataTable.Rows.Count - 1
                sSelect = oGrid.DataTable.GetValue("SELECT", oGrid.GetDataTableRowIndex(i))
                If sSelect = "Y" Then
                    v_check = True
                    Exit For
                Else
                    v_check = False
                End If
            Next

            If v_check = False Then
                sErrorMsg = "Atleast one row should be in Grid"
                Return v_check
                Exit Function
            End If

            Return v_check
        End Function

        Private Function DocumentsApproval(ByVal objForm As SAPbouiCOM.Form, ByVal sAprlStatus As String, ByRef sErrDesc As String) As Long
            Dim sFuncName As String = "DocumentsApproval"
            Dim sSelect As String = String.Empty
            Dim sEntity As String = String.Empty
            Dim sDraftKey As String = String.Empty
            Dim sWddCode As String = String.Empty
            Dim sApproverCode As String = String.Empty
            Dim sNotApproveRemarks As String = String.Empty

            Dim oApprovalRequestsService As SAPbobsCOM.ApprovalRequestsService = Nothing
            Dim oApprovalRequestParams As SAPbobsCOM.ApprovalRequestParams = Nothing
            Dim oApprovalRequest As SAPbobsCOM.ApprovalRequest = Nothing
            Dim oApprovalRequestDecision As SAPbobsCOM.ApprovalRequestDecision = Nothing

            Try
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

                oGrid = objForm.Items.Item("23").Specific
                For i = 0 To oGrid.DataTable.Rows.Count - 1
                    sSelect = oGrid.DataTable.GetValue("SELECT", oGrid.GetDataTableRowIndex(i))
                    If sSelect = "Y" Then
                        sEntity = oGrid.DataTable.GetValue("ENTITY", oGrid.GetDataTableRowIndex(i))
                        sDraftKey = oGrid.DataTable.GetValue("DRAFT NO", oGrid.GetDataTableRowIndex(i))
                        If sEntity = p_oDICompany.CompanyDB Then
                            Dim oCmpSrv As SAPbobsCOM.CompanyService = p_oDICompany.GetCompanyService

                            oApprovalRequestsService = DirectCast(oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.ApprovalRequestsService), SAPbobsCOM.ApprovalRequestsService)
                            oApprovalRequestParams = DirectCast(oApprovalRequestsService.GetDataInterface(SAPbobsCOM.ApprovalRequestsServiceDataInterfaces.arsApprovalRequestParams), SAPbobsCOM.ApprovalRequestParams)

                            Dim wddCode As Integer = oGrid.DataTable.GetValue("APPROVALCODE", oGrid.GetDataTableRowIndex(i))
                            oApprovalRequestParams.Code = wddCode
                            oApprovalRequest = oApprovalRequestsService.GetApprovalRequest(oApprovalRequestParams)
                            oApprovalRequestDecision = oApprovalRequest.ApprovalRequestDecisions.Add()
                            If sAprlStatus = "APPROVED" Then
                                oApprovalRequestDecision.Status = SAPbobsCOM.BoApprovalRequestDecisionEnum.ardApproved
                            Else
                                sNotApproveRemarks = oGrid.DataTable.GetValue("REASON FOR NOT APPROVING", oGrid.GetDataTableRowIndex(i))
                                If sNotApproveRemarks.Trim() = "" Then
                                    'sErrDesc = "Enter the reason for not approving the document"
                                    'Throw New ArgumentException(sErrDesc)
                                    Dim iRetCode As Integer
                                    sErrDesc = "Please key in the reason for ""not approving"" this PO on the extreme right hand column of this screen"
                                    iRetCode = p_oSBOApplication.MessageBox(sErrDesc, 1, "Ok")
                                   Throw New ArgumentException(sErrDesc)
                                End If
                                oApprovalRequestDecision.Status = SAPbobsCOM.BoApprovalRequestDecisionEnum.ardNotApproved
                            End If
                            oApprovalRequestDecision.Remarks = oGrid.DataTable.GetValue("REASON FOR NOT APPROVING", oGrid.GetDataTableRowIndex(i))
                            ''      oApprovalRequestDecision.ApproverUserName = SAPB1UserName
                            ''    oApprovalRequestDecision.ApproverPassword = SAPB1Password

                            oApprovalRequestsService.UpdateRequest(oApprovalRequest)

                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling EmailToUser()", sFuncName)
                            If EmailToUser(objForm, p_oDICompany, sAprlStatus, i, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                        Else
                            If ConnectToTargetCompany(p_oTargetCompany, sEntity, sApproverCode, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            If p_oTargetCompany.Connected Then
                                Dim oCmpSrv As SAPbobsCOM.CompanyService = p_oTargetCompany.GetCompanyService

                                oApprovalRequestsService = DirectCast(oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.ApprovalRequestsService), SAPbobsCOM.ApprovalRequestsService)
                                oApprovalRequestParams = DirectCast(oApprovalRequestsService.GetDataInterface(SAPbobsCOM.ApprovalRequestsServiceDataInterfaces.arsApprovalRequestParams), SAPbobsCOM.ApprovalRequestParams)

                                Dim wddCode As Integer = oGrid.DataTable.GetValue("APPROVALCODE", oGrid.GetDataTableRowIndex(i))
                                oApprovalRequestParams.Code = wddCode
                                oApprovalRequest = oApprovalRequestsService.GetApprovalRequest(oApprovalRequestParams)
                                oApprovalRequestDecision = oApprovalRequest.ApprovalRequestDecisions.Add()
                                If sAprlStatus = "APPROVED" Then
                                    oApprovalRequestDecision.Status = SAPbobsCOM.BoApprovalRequestDecisionEnum.ardApproved
                                Else
                                    sNotApproveRemarks = oGrid.DataTable.GetValue("REASON FOR NOT APPROVING", oGrid.GetDataTableRowIndex(i))
                                    If sNotApproveRemarks.Trim() = "" Then
                                        'sErrDesc = "Enter the reason for not approving the document"
                                        'Throw New ArgumentException(sErrDesc)
                                        sErrDesc = "Please key in the reason for ""not approving"" this PO on the extreme right hand column of this screen"
                                        Dim iRetCode As Integer
                                        iRetCode = p_oSBOApplication.MessageBox(sErrDesc, 1, "Ok")
                                        Throw New ArgumentException(sErrDesc)
                                    End If
                                    oApprovalRequestDecision.Status = SAPbobsCOM.BoApprovalRequestDecisionEnum.ardNotApproved
                                End If
                                oApprovalRequestDecision.Remarks = oGrid.DataTable.GetValue("REASON FOR NOT APPROVING", oGrid.GetDataTableRowIndex(i))

                                ''      oApprovalRequestDecision.ApproverUserName = SAPB1UserName
                                ''    oApprovalRequestDecision.ApproverPassword = SAPB1Password
                                oApprovalRequestsService.UpdateRequest(oApprovalRequest)

                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling EmailToUser()", sFuncName)
                                If EmailToUser(objForm, p_oTargetCompany, sAprlStatus, i, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                                p_oTargetCompany.Disconnect()
                            Else
                                p_oSBOApplication.StatusBar.SetText("Target company Not Connected", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End If
                        End If
                    End If
                Next

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
                DocumentsApproval = RTN_SUCCESS
            Catch ex As Exception
                sErrDesc = "Error while approving the Draft No. " & sDraftKey & " Entity : " & sEntity
                sErrDesc = sErrDesc & " " & ex.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                DocumentsApproval = RTN_ERROR
            End Try

        End Function

        Private Function EmailToUser(ByVal objForm As SAPbouiCOM.Form, ByVal oCompany As SAPbobsCOM.Company, ByVal sAprlStatus As String, ByVal iLine As Integer, ByRef sErrDesc As String) As Long
            Dim sFuncName As String = "EmailToUser"
            Dim sBody As String = String.Empty
            Dim p_SyncDateTime As String = String.Empty
            Dim sEmailSubject As String = String.Empty
            Dim sUSerName As String = String.Empty
            Dim sUser As String = String.Empty
            Dim sEntity As String = String.Empty
            Dim sDraftKey As String = String.Empty
            Dim sQuery As String = String.Empty
            Dim sDocType As String = String.Empty
            Dim sRemarks As String = String.Empty

            Try
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

                oEdit = objForm.Items.Item("8").Specific
                sUSerName = oEdit.Value

                oGrid = objForm.Items.Item("23").Specific

                sDocType = oGrid.DataTable.GetValue("DOCUMENT TYPE", oGrid.GetDataTableRowIndex(iLine))
                sEntity = oGrid.DataTable.GetValue("ENTITY", oGrid.GetDataTableRowIndex(iLine))
                sDraftKey = oGrid.DataTable.GetValue("DRAFT NO", oGrid.GetDataTableRowIndex(iLine))
                sRemarks = oGrid.DataTable.GetValue("REASON FOR NOT APPROVING", oGrid.GetDataTableRowIndex(iLine))
                sUser = "%/" & oCompany.UserName & "/%"

                If sDocType = "PURCHASE ORDER" Then
                    If sAprlStatus = "APPROVED" Then
                        sQuery = "update " & sHoldingDB & " ..[AB_EmailStatus] set [Status] = 'Open' where seq = " & _
                                 " (select top (1) seq + 1 from " & sHoldingDB & " ..[AB_EmailStatus] where [sUser] like '" & sUser & "' and draftkey = '" & sDraftKey & "' " & _
                                 " and DocType = 'PO' and Entity = '" & oCompany.CompanyName & "')  and draftkey = '" & sDraftKey & "' and DocType = 'PO' " & _
                                 " and Entity = '" & oCompany.CompanyName & "' and  [Status] = 'Pending' "
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Update Next level " & sQuery, sFuncName)
                        oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRecordSet.DoQuery(sQuery)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
                    Else
                        sEmailSubject = "PO Draft No. " & sDraftKey & "  " & oCompany.CompanyName & " has been Rejected "
                        p_SyncDateTime = Format(Now, "dddd") & ", " & Format(Now, "MMM") & " " & Format(Now, "dd") & ", " & Format(Now, "yyyy") & " " & Format(Now, "HH:mm:ss")
                        sBody = sBody & "<div align=left style='font-size:10.0pt;font-family:Arial'>"
                        sBody = sBody & " Dear Sir/Madam,<br /><br />"
                        sBody = sBody & p_SyncDateTime & " <br /><br />"
                        sBody = sBody & " " & " <B> Rejected your PO approval in SAP . </B><br /><br />"
                        sBody = sBody & " " & "<B> PO Draft No. : " & sDraftKey & " </B> (Can be viewed under Main Menu/ Administration/ Approval Procedures/ Approval Status Report) <br />"

                        sBody = sBody & " " & " Doc Rejected by : " & sUSerName & " <br />"
                        sBody = sBody & " " & " Entity          : " & oCompany.CompanyName
                        sBody = sBody & " " & " Remarks        : " & sRemarks.Trim()
                        sBody = sBody & "<br /><br />"
                        sBody = sBody & "Thank you."
                        sBody = sBody & "<br /><br />"
                        sBody = sBody & " Please do not reply to this email. <div/>"

                        sQuery = "UPDATE " & sHoldingDB & " ..[AB_EmailStatus] SET [Status] = 'Closed' WHERE  " & _
                                 "   draftkey = '" & sDraftKey & "' AND DocType = 'PO' AND Entity = '" & oCompany.CompanyName & "'"
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Changing Status to Closed " & sQuery, sFuncName)
                        oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRecordSet.DoQuery(sQuery)

                        sQuery = "UPDATE " & sHoldingDB & " ..[AB_EmailStatus] SET [Status] = 'Open', [EmailBody] = '" & Replace(sBody, "'", "''") & "', [EmailSub] = '" & sEmailSubject & "' WHERE seq = " & _
                           "(SELECT TOP (1) seq FROM " & sHoldingDB & " ..[AB_EmailStatus] WHERE draftkey = '" & sDraftKey & "' and DocType = 'PO' and Entity = '" & oCompany.CompanyName & "' order by cast(Seq as integer) Desc)  and draftkey = '" & sDraftKey & "' and DocType = 'PO' and Entity = '" & oCompany.CompanyName & "'"
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Triggering to Originator " & sQuery, sFuncName)
                        oRecordSet.DoQuery(sQuery)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)

                    End If
                ElseIf sDocType = "PURCHASE REQUEST" Then
                    If sAprlStatus = "APPROVED" Then
                        sQuery = "update " & sHoldingDB & " ..[AB_EmailStatus] set [Status] = 'Open' where seq = " & _
                            "(select top (1) seq + 1 from " & sHoldingDB & " ..[AB_EmailStatus] where [sUser] like '" & sUser & "' and draftkey = '" & sDraftKey & "' and DocType = 'PR' and Entity = '" & oCompany.CompanyName & "') and  draftkey = '" & sDraftKey & "' and DocType = 'PR' and Entity = '" & oCompany.CompanyName & "' and  [Status]='Pending'"
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Update Next level " & sQuery, sFuncName)
                        oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRecordSet.DoQuery(sQuery)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
                    Else
                        sEmailSubject = "PR Draft No. " & sDraftKey & "  " & oCompany.CompanyName & " has been Rejected "
                        p_SyncDateTime = Format(Now, "dddd") & ", " & Format(Now, "MMM") & " " & Format(Now, "dd") & ", " & Format(Now, "yyyy") & " " & Format(Now, "HH:mm:ss")
                        sBody = sBody & "<div align=left style='font-size:10.0pt;font-family:Arial'>"
                        sBody = sBody & " Dear Sir/Madam,<br /><br />"
                        sBody = sBody & p_SyncDateTime & " <br /><br />"
                        sBody = sBody & " " & " <B> Rejected your PR approval in SAP . </B><br /><br />"
                        sBody = sBody & " " & "<B> PR Draft No. : " & sDraftKey & " </B> (Can be viewed under Main Menu/ Administration/ Approval Procedures/ Approval Status Report) <br />"
                        sBody = sBody & " " & " Doc Rejected by : " & oCompany.UserName & " <br />"
                        sBody = sBody & " " & " Entity          : " & oCompany.CompanyName
                        sBody = sBody & " " & " Remarks         : " & sRemarks.Trim()
                        sBody = sBody & "<br /><br />"
                        sBody = sBody & "Thank you."
                        sBody = sBody & "<br /><br />"
                        sBody = sBody & " Please do not reply to this email. <div/>"

                        sQuery = "update " & sHoldingDB & " ..[AB_EmailStatus] set [Status] = 'Closed' where " & _
                                 "   draftkey = '" & sDraftKey & "' and DocType = 'PR' and Entity = '" & oCompany.CompanyName & "'"
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Changing Status to Closed " & sQuery, sFuncName)
                        oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRecordSet.DoQuery(sQuery)

                        sQuery = "update " & sHoldingDB & " ..[AB_EmailStatus] set [Status] = 'Open', [EmailBody] = '" & Replace(sBody, "'", "''") & "', [EmailSub] = '" & sEmailSubject & "' where seq = " & _
                           "(select top (1) seq  from " & sHoldingDB & " ..[AB_EmailStatus] where draftkey = '" & sDraftKey & "' and DocType = 'PR' and Entity = '" & oCompany.CompanyName & "'  order by cast(Seq as integer) Desc)  and draftkey = '" & sDraftKey & "' and DocType = 'PR' and Entity = '" & oCompany.CompanyName & "'"
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Triggering to Originator " & sQuery, sFuncName)
                        oRecordSet.DoQuery(sQuery)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
                    End If
                End If
               

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
                EmailToUser = RTN_SUCCESS
            Catch ex As Exception
                sErrDesc = ex.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                EmailToUser = RTN_ERROR
            End Try
        End Function

        Public Sub ApprovalWindow_SBO_ItemEvent(ByVal FormUID As String, ByVal pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean, ByVal objForm As SAPbouiCOM.Form)
            Dim sFuncName As String = "ApprovalWindow_SBO_ItemEvent"
            Dim sErrDesc As String = String.Empty

            Try
                If pval.Before_Action = True Then
                    Select Case pval.EventType
                        Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                            objForm = p_oSBOApplication.Forms.GetForm(pval.FormTypeEx, pval.FormTypeCount)

                        Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                            objForm = p_oSBOApplication.Forms.GetForm(pval.FormTypeEx, pval.FormTypeCount)
                            If pval.ItemUID = "24" Then
                                'oEdit = objForm.Items.Item("18").Specific
                                'If oEdit.Value = "" Then
                                '    p_oSBOApplication.StatusBar.SetText("Select Date From", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                '    BubbleEvent = False
                                '    Exit Sub
                                'End If
                                oEdit = objForm.Items.Item("20").Specific
                                If oEdit.Value = "" Then
                                    p_oSBOApplication.StatusBar.SetText("Select Date To", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            ElseIf pval.ItemUID = "3" Then
                                If CheckFields(objForm) = False Then
                                    p_oSBOApplication.StatusBar.SetText(sErrorMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                    Exit Sub
                                Else
                                    p_oSBOApplication.StatusBar.SetText("Processing.. Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    If DocumentsApproval(objForm, "APPROVED", sErrDesc) = RTN_SUCCESS Then
                                        objForm.Freeze(True)
                                        LoadGrid(objForm)
                                        objForm.Freeze(False)
                                        objForm.Update()
                                    Else
                                        Throw New ArgumentException(sErrDesc)
                                    End If
                                    p_oSBOApplication.StatusBar.SetText("Operation Completed successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                End If
                            ElseIf pval.ItemUID = "4" Then
                                If CheckFields(objForm) = False Then
                                    p_oSBOApplication.StatusBar.SetText(sErrorMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                    Exit Sub
                                Else
                                    p_oSBOApplication.StatusBar.SetText("Processing.. Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    If DocumentsApproval(objForm, "REJECTED", sErrDesc) = RTN_SUCCESS Then
                                        objForm.Freeze(True)
                                        LoadGrid(objForm)
                                        objForm.Freeze(False)
                                        objForm.Update()
                                    Else
                                        Throw New ArgumentException(sErrDesc)
                                    End If
                                    p_oSBOApplication.StatusBar.SetText("Operation Completed successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                End If
                            End If

                        Case SAPbouiCOM.BoEventTypes.et_VALIDATE And pval.InnerEvent = False
                            objForm = p_oSBOApplication.Forms.GetForm(pval.FormTypeEx, pval.FormTypeCount)
                            If pval.ItemUID = "22" Then
                                Dim orset As SAPbobsCOM.Recordset = Nothing
                                Dim sSQL As String = String.Empty
                                If Not String.IsNullOrEmpty(objForm.Items.Item("22").Specific.String) Then
                                    sSQL = "SELECT T0.[U_AB_COMPANYNAME] FROM " & sHoldingDB & " ..[@AB_COMPANYDATA]  T0 WHERE T0.[Name]  = '" & objForm.Items.Item("22").Specific.String & "'"
                                    orset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    orset.DoQuery(sSQL)
                                    objForm.Items.Item("1000001").Specific.String = orset.Fields.Item("U_AB_COMPANYNAME").Value

                                End If

                                ''1000001

                            End If

                        Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                            objForm = p_oSBOApplication.Forms.GetForm(pval.FormTypeEx, pval.FormTypeCount)
                            If pval.ItemUID = "23" Then
                                oGrid = objForm.Items.Item("23").Specific
                                If pval.ColUID = "DRAFT NO" Then
                                    Dim sDraftNo, sEntity, sDocType, sApprovedBy As String
                                    sDraftNo = oGrid.DataTable.GetValue("DRAFT NO", oGrid.GetDataTableRowIndex(pval.Row))
                                    If sDraftNo <> "" Then
                                        sEntity = oGrid.DataTable.GetValue("ENTITY", oGrid.GetDataTableRowIndex(pval.Row))
                                        sDocType = oGrid.DataTable.GetValue("DOCUMENT TYPE", oGrid.GetDataTableRowIndex(pval.Row))
                                        sApprovedBy = oGrid.DataTable.GetValue("APPROVED BY", oGrid.GetDataTableRowIndex(pval.Row))
                                        If sApprovedBy = "You Are The First Approver" Then
                                            sApprovedBy = ""
                                        End If
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Entity is " & sEntity & " and Connected Company is " & p_oDICompany.CompanyDB, sFuncName)
                                        If sEntity.ToUpper() <> p_oDICompany.CompanyDB.ToUpper() Then
                                            BubbleEvent = False
                                            If sDocType = "PURCHASE ORDER" Then
                                                InitializePRPOForm(sEntity, sDraftNo, "PO", sApprovedBy)
                                            ElseIf sDocType = "PURCHASE REQUEST" Then
                                                InitializePRPOForm(sEntity, sDraftNo, "PR", sApprovedBy)
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                    End Select
                Else
                    Select Case pval.EventType
                        Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                            objForm = p_oSBOApplication.Forms.GetForm(pval.FormTypeEx, pval.FormTypeCount)
                            If pval.ItemUID = "24" Then
                                p_oSBOApplication.StatusBar.SetText("Processing.. Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                objForm.Freeze(True)
                                LoadGrid(objForm)
                                objForm.Freeze(False)
                                p_oSBOApplication.StatusBar.SetText("Operation Completed successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            ElseIf pval.ItemUID = "23" Then
                                oGrid = objForm.Items.Item("23").Specific
                                If pval.Row > -1 Then
                                    If pval.ColUID = "SELECT" Then
                                        Dim sSelectColValue As String
                                        sSelectColValue = oGrid.DataTable.GetValue("SELECT", oGrid.GetDataTableRowIndex(pval.Row))
                                        If sSelectColValue = "Y" Then
                                            oGrid.CommonSetting.SetRowBackColor(pval.Row + 1, RGB(255, 255, 0))
                                        Else
                                            oGrid.CommonSetting.SetRowBackColor(pval.Row + 1, RGB(231, 231, 231))
                                        End If
                                    End If
                                End If
                            End If

                        Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                            objForm = p_oSBOApplication.Forms.GetForm(pval.FormTypeEx, pval.FormTypeCount)
                            If pval.ItemUID = "23" Then
                                oGrid = objForm.Items.Item("23").Specific
                                If pval.ColUID = "SELECT" Then
                                    If oGrid.DataTable.Rows.Count - 1 > -1 Then
                                        p_oSBOApplication.StatusBar.SetText("Processing...Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        objForm.Freeze(True)
                                        If bdoubleCheck = False Then
                                            For i As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                                                oGrid.DataTable.SetValue("SELECT", i, "Y")
                                            Next
                                            bdoubleCheck = True
                                        ElseIf bdoubleCheck = True Then
                                            For i As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                                                oGrid.DataTable.SetValue("SELECT", i, "N")
                                            Next
                                            bdoubleCheck = False
                                        End If
                                        objForm.Freeze(False)
                                        p_oSBOApplication.StatusBar.SetText("Operation completed successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                                    End If
                                End If
                            End If

                        Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                            ''objForm = p_oSBOApplication.Forms.GetForm(pval.FormTypeEx, pval.FormTypeCount)
                            ''Try
                            ''    Dim objItem, oItem, oItem1 As SAPbouiCOM.Item

                            ''    'Second Line
                            ''    objItem = objForm.Items.Item("13")
                            ''    oItem = objForm.Items.Item("12")
                            ''    objItem.Top = oItem.Top
                            ''    objItem.Left = oItem.Width + oItem.Left + 5

                            ''    objItem = objForm.Items.Item("14")
                            ''    oItem = objForm.Items.Item("13")
                            ''    objItem.Top = oItem.Top
                            ''    objItem.Left = oItem.Width + oItem.Left + 5

                            ''    objItem = objForm.Items.Item("9")
                            ''    oItem = objForm.Items.Item("14")
                            ''    objItem.Top = oItem.Top
                            ''    oItem1 = objForm.Items.Item("25")
                            ''    objItem.Left = oItem.Width + oItem.Left + oItem1.Width + 25

                            ''    objItem = objForm.Items.Item("10")
                            ''    oItem = objForm.Items.Item("9")
                            ''    objItem.Top = oItem.Top
                            ''    objItem.Left = oItem.Width + oItem.Left + 5

                            ''    'First Line
                            ''    objItem = objForm.Items.Item("7")
                            ''    oItem = objForm.Items.Item("6")
                            ''    objItem.Top = oItem.Top
                            ''    objItem.Left = oItem.Width + oItem.Left + 5

                            ''    objItem = objForm.Items.Item("8")
                            ''    oItem = objForm.Items.Item("7")
                            ''    objItem.Top = oItem.Top
                            ''    oItem = objForm.Items.Item("14")
                            ''    objItem.Left = oItem.Left

                            ''    objItem = objForm.Items.Item("25")
                            ''    oItem = objForm.Items.Item("8")
                            ''    objItem.Top = oItem.Top
                            ''    objItem.Left = oItem.Left + oItem.Width + 5

                            ''    objItem = objForm.Items.Item("21")
                            ''    oItem = objForm.Items.Item("25")
                            ''    objItem.Top = oItem.Top
                            ''    oItem = objForm.Items.Item("9")
                            ''    objItem.Left = oItem.Left

                            ''    objItem = objForm.Items.Item("22")
                            ''    oItem = objForm.Items.Item("21")
                            ''    objItem.Top = oItem.Top
                            ''    oItem = objForm.Items.Item("10")
                            ''    objItem.Left = oItem.Left

                            ''    'third line
                            ''    objItem = objForm.Items.Item("19")
                            ''    oItem = objForm.Items.Item("18")
                            ''    objItem.Top = oItem.Top
                            ''    oItem = objForm.Items.Item("13")
                            ''    objItem.Left = oItem.Left

                            ''    objItem = objForm.Items.Item("20")
                            ''    oItem = objForm.Items.Item("19")
                            ''    objItem.Top = oItem.Top
                            ''    oItem = objForm.Items.Item("14")
                            ''    objItem.Left = oItem.Left

                            ''    objItem = objForm.Items.Item("15")
                            ''    oItem = objForm.Items.Item("20")
                            ''    objItem.Top = oItem.Top
                            ''    oItem = objForm.Items.Item("9")
                            ''    objItem.Left = oItem.Left

                            ''    objItem = objForm.Items.Item("16")
                            ''    oItem = objForm.Items.Item("15")
                            ''    objItem.Top = oItem.Top
                            ''    oItem = objForm.Items.Item("10")
                            ''    objItem.Left = oItem.Left

                            ''Catch ex As Exception
                            ''    objForm.Freeze(False)
                            ''    objForm.Update()
                            ''End Try

                        Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                            objForm = p_oSBOApplication.Forms.GetForm(pval.FormTypeEx, pval.FormTypeCount)
                            Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                            oCFLEvento = pval
                            Dim sCFL_ID As String
                            sCFL_ID = oCFLEvento.ChooseFromListUID
                            'Dim objForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.Item(FormUID)
                            Dim oCFL As SAPbouiCOM.ChooseFromList
                            oCFL = objForm.ChooseFromLists.Item(sCFL_ID)
                            Try
                                If oCFLEvento.BeforeAction = False Then
                                    Dim oDataTable As SAPbouiCOM.DataTable
                                    oDataTable = oCFLEvento.SelectedObjects
                                    If pval.ItemUID = "10" Then
                                        objForm.Items.Item("10").Specific.string = oDataTable.GetValue("U_NAME", 0)
                                    ElseIf pval.ItemUID = "16" Then
                                        objForm.Items.Item("16").Specific.string = oDataTable.GetValue("CardName", 0)
                                    End If
                                End If
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

        Public Sub Approvalwindow_SBO_MenuEvent(ByVal pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Dim sFuncName As String = "Approvalwindow_SBO_MenuEvent"
            Dim sErrDesc As String = String.Empty

            Try
                If pVal.BeforeAction = False Then
                    Dim objForm As SAPbouiCOM.Form
                    If pVal.MenuUID = "APRL" Then
                        InitializeForm()
                    End If
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


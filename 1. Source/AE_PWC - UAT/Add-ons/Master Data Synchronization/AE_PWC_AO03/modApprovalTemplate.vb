Imports System.Xml
Imports System.IO
Imports SAPbobsCOM

Namespace AE_PWC_AO03

    Module modApprovalTemplate

        Dim sPath As String = String.Empty
        Dim sFuncName As String = String.Empty
        Dim ival As Integer
        Dim IsError As Boolean
        Dim iErr As Integer = 0
        Dim sErr As String = String.Empty
        Dim xDoc As New XmlDocument

        Dim oMatrix As SAPbouiCOM.Matrix = Nothing

        ''Public Function ApprovalTemplateSync(ByRef oHoldingCompany As SAPbobsCOM.Company, ByRef oTragetCompany As SAPbobsCOM.Company, ByVal sMasterdatacode As String, _
        ''                            ByRef sErrDesc As String) As Long

        ''    'Function   :   ApprovalTemplateSync()
        ''    'Purpose    :   
        ''    'Parameters :   ByVal oForm As SAPbouiCOM.Form
        ''    '                   oForm=Form Type
        ''    '               ByRef sErrDesc As String
        ''    '                   sErrDesc=Error Description to be returned to calling function
        ''    '               
        ''    '                   =
        ''    'Return     :   0 - FAILURE
        ''    '               1 - SUCCESS
        ''    'Author     :   SRI
        ''    'Date       :   30/12/2007
        ''    'Change     :

        ''    Dim sFuncName As String = String.Empty
        ''    Dim iStagingcode As Integer
        ''    Dim sSQLString As String = String.Empty

        ''    Dim oRset_Tar As SAPbobsCOM.Recordset = Nothing
        ''    Dim oRset_Hol As SAPbobsCOM.Recordset = Nothing

        ''    Try
        ''        sFuncName = "BPMaterSync()"
        ''        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

        ''        oRset_Hol = oHoldingCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
        ''        oRset_Tar = oTragetCompany.GetBusinessObject(BoObjectTypes.BoRecordset)

        ''        Dim oCmpSrv As SAPbobsCOM.CompanyService
        ''        Dim oApprovalStagesService As ApprovalStagesService
        ''        Dim oApprovalStage As ApprovalStage = Nothing
        ''        Dim oApprovalStageApprovers As ApprovalStageApprovers
        ''        Dim oApprover As ApprovalStageApprover
        ''        Dim oApprovalStageParams As ApprovalStageParams
        ''        ''Template
        ''        Dim oApprovalTemplateStage As ApprovalTemplateStage
        ''        Dim oApprovalTemplate As ApprovalTemplate
        ''        Dim oApprovalTemplateParams As ApprovalTemplateParams
        ''        Dim oApprovalTemplateTerm As ApprovalTemplateTerm
        ''        Dim oApprovalTemplateService As ApprovalTemplatesService

        ''        sSQLString = "SELECT T0.[WstCode], T0.[Name], T0.[Remarks], T0.[MaxReqr], T0.[MaxRejReqr], " & _
        ''        " T1.[Userid] FROM OWST T0  INNER JOIN WST1 T1 ON T0.[WstCode] = T1.[WstCode] WHERE T0.[WstCode] = (SELECT T0.[WstCode] FROM WTM2 T0 WHERE T0.[WtmCode]  = '" & sMasterdatacode & "')"
        ''        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Get the staging code " & sSQLString, sFuncName)
        ''        oRset_Hol.DoQuery(sSQLString)
        ''        ' ''Stages
        ''        oApprovalStagesService = oCmpSrv.GetBusinessService(ServiceTypes.ApprovalStagesService)
        ''        'get new Approval Stage
        ''        oApprovalStage = oApprovalStagesService.GetDataInterface(ApprovalStagesServiceDataInterfaces.assdiApprovalStage)
        ''        oApprovalTemplate = oApprovalTemplateService.GetDataInterface(ApprovalTemplatesServiceDataInterfaces.atsdiApprovalTemplate)
        ''        'set the name
        ''        oApprovalStage.Name = oRset_Hol.Fields.Item("Name").Value
        ''        oApprovalStage.Remarks = oRset_Hol.Fields.Item("Remarks").Value
        ''        'get ApprovalStageApprovers collection
        ''        oApprovalStageApprovers = oApprovalStage.ApprovalStageApprovers
        ''        'add new Approver
        ''        For imjs As Integer = 1 To oRset_Hol.RecordCount
        ''            oApprover = oApprovalStageApprovers.Add
        ''            oApprover.UserID = oRset_Hol.Fields.Item("Userid").Value
        ''        Next
        ''        'set the number of required approvers
        ''        oApprovalStage.NoOfApproversRequired = oRset_Hol.Fields.Item("MaxReqr").Value
        ''        'add Approval Stage
        ''        oApprovalStageParams = oApprovalStagesService.AddApprovalStage(oApprovalStage)

        ''        sSQLString = "SELECT T0.[WstCode], T0.[Name], T0.[Remarks], T0.[MaxReqr], T0.[MaxRejReqr], " & _
        ''        " T1.[Userid] FROM OWST T0  INNER JOIN WST1 T1 ON T0.[WstCode] = T1.[WstCode] WHERE T0.[WstCode] = (SELECT T0.[WstCode] FROM WTM2 T0 WHERE T0.[WtmCode]  = '" & sMasterdatacode & "')"
        ''        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Get the Approval Template code " & sSQLString, sFuncName)
        ''        oRset_Hol.DoQuery(sSQLString)


        ''        'set the name of the Approval Template
        ''        oApprovalTemplate.Name = txtName.Text

        ''        'add the user that need the approval(userId=3 is "Fred")
        ''        oApprovalTemplate.ApprovalTemplateUsers.Add.UserID = txtUserCode.Text

        ''        'Add the checked documnets  
        ''        Call AddCheckedDocumnets(oApprovalTemplate)

        ''        'get Approval Stages
        ''        oApprovalTemplateStage = oApprovalTemplate.ApprovalTemplateStages.Add

        ''        'set the code of an existing stage(e.g code=1 the stage name is Accounting)
        ''        oApprovalTemplateStage.ApprovalStageCode = CInt(txtStageCode.Text)

        ''        'include terms in the template 
        ''        oApprovalTemplate.UseTerms = IIf(chkIncludeTerms.Checked = True, BoYesNoEnum.tYES, BoYesNoEnum.tNO)

        ''        ' ''add new term
        ''        ''oApprovalTemplateTerm = oApprovalTemplate.ApprovalTemplateTerms.Add

        ''        ' ''set the condition
        ''        ''oApprovalTemplateTerm.ConditionType = GetCondition(cboConditions.SelectedItem)

        ''        ' ''set the Operation Type
        ''        ''oApprovalTemplateTerm.OperationType = GetOperation(cboOperation.SelectedItem)


        ''        ' ''set the value
        ''        ''oApprovalTemplateTerm.Value = txtTermValue.Text

        ''        'oApprovalTemplateTerm = oApprovalTemplate.ApprovalTemplateQueries.Add
        ''        oApprovalTemplate.ApprovalTemplateQueries.Add.QueryID = 298

        ''        'add Approval Template
        ''        oApprovalTemplateParams = oApprovalTemplateService.AddApprovalTemplate(oApprovalTemplate)







        ''        BPMaterSync = RTN_SUCCESS
        ''    Catch ex As Exception
        ''        BPMaterSync = RTN_ERROR
        ''        sErrDesc = ex.Message
        ''        Call WriteToLogFile(sErrDesc, sFuncName)
        ''        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        ''    Finally
        ''        System.Runtime.InteropServices.Marshal.ReleaseComObject(oBP_Holding)
        ''        System.Runtime.InteropServices.Marshal.ReleaseComObject(oBP_Target)
        ''        System.Runtime.InteropServices.Marshal.ReleaseComObject(oBP_Holding_Banks)
        ''        System.Runtime.InteropServices.Marshal.ReleaseComObject(oBP_Target_Banks)
        ''        oBP_Holding = Nothing
        ''        oBP_Target = Nothing
        ''        oRset_Tar = Nothing
        ''        oRset_Hol = Nothing
        ''        oDlfPaymenthod = Nothing

        ''        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Releasing the Objects", sFuncName)
        ''    End Try

        ''End Function

    End Module

End Namespace



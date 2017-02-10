Option Explicit On
'Imports SAPbouiCOM.Framework
Imports System.Windows.Forms


Public Class clsEventHandler
    Dim WithEvents SBO_Application As SAPbouiCOM.Application ' holds connection with SBO
    Dim p_oDICompany As New SAPbobsCOM.Company

    Public Sub New(ByRef oApplication As SAPbouiCOM.Application, ByRef oCompany As SAPbobsCOM.Company)
        Dim sFuncName As String = String.Empty
        Try
            sFuncName = "Class_Initialize()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Retriving SBO Application handle", sFuncName)
            SBO_Application = oApplication
            p_oDICompany = oCompany

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch exc As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Call WriteToLogFile(exc.Message, sFuncName)
        End Try
    End Sub

    Public Function SetApplication(ByRef sErrDesc As String) As Long
        ' **********************************************************************************
        '   Function   :    SetApplication()
        '   Purpose    :    This function will be calling to initialize the default settings
        '                   such as Retrieving the Company Default settings, Creating Menus, and
        '                   Initialize the Event Filters
        '               
        '   Parameters :    ByRef sErrDesc AS string
        '                       sErrDesc = Error Description to be returned to calling function
        '   Return     :    0 - FAILURE
        '                   1 - SUCCESS
        ' **********************************************************************************
        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "SetApplication()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling SetMenus()", sFuncName)
            If SetMenus(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling SetFilters()", sFuncName)
            If SetFilters(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            SetApplication = RTN_SUCCESS
        Catch exc As Exception
            sErrDesc = exc.Message
            Call WriteToLogFile(exc.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            SetApplication = RTN_ERROR
        End Try
    End Function

    Private Function SetMenus(ByRef sErrDesc As String) As Long
        ' **********************************************************************************
        '   Function   :    SetMenus()
        '   Purpose    :    This function will be gathering to create the customized menu
        '               
        '   Parameters :    ByRef sErrDesc AS string
        '                       sErrDesc = Error Description to be returned to calling function
        '   Return     :    0 - FAILURE
        '                   1 - SUCCESS
        ' **********************************************************************************
        Dim sFuncName As String = String.Empty
        ' Dim oMenuItem As SAPbouiCOM.MenuItem
        Try
            sFuncName = "SetMenus()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            SetMenus = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            SetMenus = RTN_ERROR
        End Try
    End Function

    Private Function SetFilters(ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function   :    SetFilters()
        '   Purpose    :    This function will be gathering to declare the event filter 
        '                   before starting the AddOn Application
        '               
        '   Parameters :    ByRef sErrDesc AS string
        '                       sErrDesc = Error Description to be returned to calling function
        '   Return     :    0 - FAILURE
        '                   1 - SUCCESS
        ' **********************************************************************************

        Dim oFilters As SAPbouiCOM.EventFilters
        Dim oFilter As SAPbouiCOM.EventFilter
        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "SetFilters()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Initializing EventFilters object", sFuncName)
            oFilters = New SAPbouiCOM.EventFilters



            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding filters", sFuncName)
            SBO_Application.SetFilter(oFilters)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            SetFilters = RTN_SUCCESS
        Catch exc As Exception
            sErrDesc = exc.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            SetFilters = RTN_ERROR
        End Try
    End Function

    Private Sub SBO_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles SBO_Application.AppEvent
        ' **********************************************************************************
        '   Function   :    SBO_Application_AppEvent()
        '   Purpose    :    This function will be handling the SAP Application Event
        '               
        '   Parameters :    ByVal EventType As SAPbouiCOM.BoAppEventTypes
        '                       EventType = set the SAP UI Application Eveny Object        
        ' **********************************************************************************
        Dim sFuncName As String = String.Empty
        Dim sErrDesc As String = String.Empty
        Dim sMessage As String = String.Empty

        Try
            sFuncName = "SBO_Application_AppEvent()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            Select Case EventType
                Case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged, SAPbouiCOM.BoAppEventTypes.aet_ShutDown, SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition
                    sMessage = String.Format("Please wait for a while to disconnect the AddOn {0} ....", System.Windows.Forms.Application.ProductName)
                    p_oSBOApplication.SetStatusBarMessage(sMessage, SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                    End
            End Select

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch ex As Exception
            sErrDesc = ex.Message
            WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            ShowErr(sErrDesc)
        Finally
            GC.Collect()  'Forces garbage collection of all generations.
        End Try
    End Sub

    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, _
            ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        ' **********************************************************************************
        '   Function   :    SBO_Application_ItemEvent()
        '   Purpose    :    This function will be handling the SAP Menu Event
        '               
        '   Parameters :    ByVal FormUID As String
        '                       FormUID = set the FormUID
        '                   ByRef pVal As SAPbouiCOM.ItemEvent
        '                       pVal = set the SAP UI ItemEvent Object
        '                   ByRef BubbleEvent As Boolean
        '                       BubbleEvent = set the True/False        
        ' **********************************************************************************

        Dim sErrDesc As String = String.Empty
        Dim sFuncName As String = String.Empty

        Dim oDTDistinct As DataTable = Nothing
        Dim oDTRowFilter As DataTable = Nothing



        Try
            'sFuncName = "SBO_Application_ItemEvent()"
            'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            If Not IsNothing(p_oDICompany) Then
                If Not p_oDICompany.Connected Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectDICompSSO()", sFuncName)
                    If ConnectDICompSSO(p_oDICompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                End If
            End If

            If pVal.BeforeAction = False Then

                Select Case pVal.FormTypeEx

                    Case "10001"

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                            Dim oform As SAPbouiCOM.Form = Nothing
                            Dim oform_udf As SAPbouiCOM.Form = Nothing
                            Try
                                If p_BPTypecount > 0 Then
                                    oform = p_oSBOApplication.Forms.GetFormByTypeAndCount(134, p_FormTypecount)
                                    oform_udf = p_oSBOApplication.Forms.GetFormByTypeAndCount(-134, p_FormTypecount)
                                    If p_oCompDef.sAuthorization = "approve" And oform_udf.Items.Item("u_ab_status").Specific.value.ToString.Trim() <> "approved" Then
                                        oform.Items.Item("btnapprove").Enabled = True
                                    Else
                                        oform.Items.Item("btnapprove").Enabled = False
                                    End If
                                    p_BPTypecount = 0
                                End If

                            Catch ex As Exception
                                p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                Call WriteToLogFile(sErrDesc, sFuncName)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("completed with error", sFuncName)
                                Exit Sub
                            End Try
                        End If
                    Case "10021", "1470010112"

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                            Dim oform As SAPbouiCOM.Form = Nothing
                            Dim iformtype As Integer = 0
                            Try
                                If p_FormTypecount > 0 Then
                                    If pVal.FormTypeEx = "10021" Then
                                        iformtype = 142
                                    Else
                                        iformtype = 1470000200
                                    End If
                                    oform = p_oSBOApplication.Forms.GetFormByTypeAndCount(iformtype, p_FormTypecount)
                                    oform.Items.Item("txtnote").Enabled = False
                                    p_FormTypecount = 0
                                End If

                            Catch ex As Exception
                                p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                Call WriteToLogFile(sErrDesc, sFuncName)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                                Exit Sub
                            End Try
                        End If

                    Case "-134"
                        If pVal.ItemChanged = True Then

                            Try
                                Dim oform As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount(134, pVal.FormTypeCount)
                                Dim oform_UDF As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount(-134, pVal.FormTypeCount)
                                Dim sStatus As String = String.Empty
                                sStatus = oform_UDF.Items.Item("U_AB_STATUS").Specific.value.ToString.Trim()

                                If p_oCompDef.sAuthorization = "APPROVE" And sStatus <> "APPROVED" Then
                                    oform.Items.Item("btnapprove").Enabled = True
                                Else
                                    oform.Items.Item("btnapprove").Enabled = False
                                End If
                            Catch ex As Exception
                                p_oSBOApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                BubbleEvent = False
                                Exit Try
                            End Try
                        End If

                    Case "134"
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD Then
                            Dim oform As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount(pVal.FormTypeEx, pVal.FormTypeCount)
                            Dim oItem As SAPbouiCOM.Item = Nothing
                            Dim oRItem As SAPbouiCOM.Item = Nothing

                            oItem = oform.Items.Add("btnapprove", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
                            oRItem = oform.Items.Item("2")
                            oItem.Height = oRItem.Height
                            oItem.Width = oRItem.Width
                            oItem.Top = oRItem.Top
                            oItem.Left = oRItem.Left + oRItem.Width + 5
                            oItem.Visible = True
                            oform.Items.Item("btnapprove").Specific.caption = "Approve"
                            If p_oCompDef.sAuthorization = "APPROVE" Then
                                oform.Items.Item("btnapprove").Enabled = True
                            Else
                                oform.Items.Item("btnapprove").Enabled = False

                            End If
                        End If

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.ItemUID = "1" Then
                            Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount(134, pVal.FormTypeCount)
                            Dim oform_UDF As SAPbouiCOM.Form = Nothing
                            Try

                                oform_UDF = p_oSBOApplication.Forms.GetFormByTypeAndCount(-134, oForm.TypeCount)
                                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    If p_oCompDef.sAuthorization = "APPROVE" And oform_UDF.Items.Item("U_AB_STATUS").Specific.value.ToString.Trim() <> "APPROVED" Then
                                        oForm.Items.Item("btnapprove").Enabled = True
                                    Else
                                        oForm.Items.Item("btnapprove").Enabled = False
                                    End If
                                End If
                            Catch ex As Exception
                                p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                Call WriteToLogFile(sErrDesc, sFuncName)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                                BubbleEvent = False
                                Exit Sub
                            End Try


                        End If

                    Case "50105"

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT And pVal.ItemUID = "5" And pVal.ColUID = "3" And pVal.ItemChanged = True Then
                            Try
                                Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount(50105, pVal.FormTypeCount)
                                Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("5").Specific
                                p_sAstatus = String.Empty
                                p_sAstatus = oMatrix.Columns.Item("3").Cells.Item(pVal.Row).Specific.selected.value
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Approval Status  " & p_sAstatus, sFuncName)

                            Catch ex As Exception
                                p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Sub
                            End Try

                        End If

                    Case "142", "1470000200"

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD Then
                            Dim oform As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount(pVal.FormTypeEx, pVal.FormTypeCount)
                            Dim oItem As SAPbouiCOM.Item = Nothing
                            Dim oRItem As SAPbouiCOM.Item = Nothing
                            Dim oR1Item As SAPbouiCOM.Item = Nothing
                            Dim sString As String = String.Empty

                            Try

                                sString = "1)        Terms & Conditions (T&C): In most instances, PwC's T&C should be used. Please consult OGC if we sign any suppliers’ agreement. For details, please refer to Finance Policies and Procedures (Procurement section). "
                                sString += vbCrLf
                                sString += "2)        Procurement of IT systems and applications: Prior to making any purchase of IT systems/ applications/ services, please consult GTS managers for Infrastructure and Personal Computing. For details, please refer to Finance Policies and Procedures (Procurement section). "
                                sString += vbCrLf
                                sString += "3)        3 quotes are required for >$10,000. If this is not met, please explain in waiver section visible at User-Defined Fields and this will escalate to next level of approval."

                                oItem = oform.Items.Add("lbl", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                                oRItem = oform.Items.Item("230")
                                oItem.Height = oRItem.Height
                                oItem.Width = oRItem.Width
                                oItem.Top = oRItem.Top + oRItem.Height + 3
                                oItem.Left = oRItem.Left
                                oItem.Visible = True
                                oform.Items.Item("lbl").Specific.caption = "Note  "


                                oItem = oform.Items.Add("txtnote", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)
                                oRItem = oform.Items.Item("222")
                                oR1Item = oform.Items.Item("16")
                                oItem.Height = 62
                                oItem.Width = oform.Width - 152
                                oItem.Top = oRItem.Top + oRItem.Height + 3
                                oItem.Left = oRItem.Left
                                oItem.Visible = True
                                oItem.Enabled = False
                                oform.Items.Item("txtnote").Specific.String = sString
                                oform.Items.Item("lbl").LinkTo = "txtnote"

                            Catch ex As Exception
                                p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Sub
                            End Try
                        End If

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_RESIZE Then
                            Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
                            Try
                                oForm.Freeze(True)
                                If oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Minimized Then
                                    oForm.Items.Item("txtnote").Width = 150
                                    oForm.Items.Item("txtnote").Top = oForm.Items.Item("222").Top + oForm.Items.Item("222").Height + 3
                                    oForm.Items.Item("lbl").LinkTo = "txtnote"
                                ElseIf oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized Then
                                    oForm.Items.Item("txtnote").Width = oForm.Width - 152
                                    oForm.Items.Item("txtnote").Top = oForm.Items.Item("222").Top + oForm.Items.Item("222").Height + 3
                                    oForm.Items.Item("lbl").LinkTo = "txtnote"
                                ElseIf oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Restore Then
                                    If oForm.Items.Item("txtnote").Width = 150 Then
                                        oForm.Items.Item("txtnote").Width = oForm.Width - 152
                                        oForm.Items.Item("txtnote").Top = oForm.Items.Item("222").Top + oForm.Items.Item("222").Height + 3
                                        oForm.Items.Item("lbl").LinkTo = "txtnote"
                                    Else
                                        oForm.Items.Item("txtnote").Width = 150
                                        oForm.Items.Item("txtnote").Top = oForm.Items.Item("222").Top + oForm.Items.Item("222").Height + 3
                                        oForm.Items.Item("lbl").LinkTo = "txtnote"
                                    End If
                                End If
                                oForm.Freeze(False)
                            Catch ex As Exception
                                oForm.Freeze(False)
                                Call WriteToLogFile(sErrDesc, sFuncName)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                                p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Sub
                            End Try
                        End If

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.ItemUID = "1" Then
                            Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
                            Try
                                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    oForm.Items.Item("txtnote").Enabled = False
                                End If
                            Catch ex As Exception
                                p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                Call WriteToLogFile(sErrDesc, sFuncName)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                                BubbleEvent = False
                                Exit Sub
                            End Try


                        End If

                End Select
            Else   ' Before action True

                Select Case pVal.FormTypeEx

                    Case "10001"

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then

                            Dim oForm As SAPbouiCOM.Form = Nothing
                            Dim oForm_L As SAPbouiCOM.Form = Nothing
                            Dim oform_UDF As SAPbouiCOM.Form = Nothing
                            Try
                                oForm_L = p_oSBOApplication.Forms.ActiveForm
                                If oForm_L.Title.ToUpper = "LIST OF BUSINESS PARTNERS" And p_BPTypecount > 0 Then
                                    oForm = p_oSBOApplication.Forms.GetFormByTypeAndCount(134, p_BPTypecount)
                                    oform_UDF = p_oSBOApplication.Forms.GetFormByTypeAndCount(-134, p_BPTypecount)
                                    If p_oCompDef.sAuthorization = "APPROVE" And oform_UDF.Items.Item("U_AB_STATUS").Specific.value.ToString.Trim() <> "APPROVED" Then
                                        oForm.Items.Item("btnapprove").Enabled = True
                                    Else
                                        oForm.Items.Item("btnapprove").Enabled = False
                                    End If
                                    p_BPTypecount = 0
                                End If

                            Catch ex As Exception
                                p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Sub
                            End Try
                        End If

                    Case "134" ' BP Approval

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.ItemUID = "btnapprove" Then
                            Dim oform As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount(pVal.FormTypeEx, pVal.FormTypeCount)
                            '' Dim oform_UDF As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount(-134, pVal.FormTypeCount)
                            ''oform.Items.Item("10002044").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            Dim sBPcode As String = String.Empty
                            Dim oBP As SAPbobsCOM.BusinessPartners = Nothing
                            Dim irlt As Integer = 0
                            Try
                                If oform.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE And p_oCompDef.sAuthorization = "APPROVE" Then
                                    oBP = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
                                    sBPcode = oform.Items.Item("5").Specific.String
                                    If oBP.GetByKey(sBPcode) Then
                                        oBP.Frozen = SAPbobsCOM.BoYesNoEnum.tNO
                                        oBP.Valid = SAPbobsCOM.BoYesNoEnum.tYES
                                        oBP.UserFields.Fields.Item("U_AB_STATUS").Value = "APPROVED"
                                        irlt = oBP.Update()
                                        If irlt <> 0 Then
                                            p_oSBOApplication.SetStatusBarMessage(p_oDICompany.GetLastErrorCode & " - " & p_oDICompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                            BubbleEvent = False
                                            Exit Try
                                        Else
                                            oform.Close()
                                            p_oSBOApplication.ActivateMenuItem("2561")
                                            oform = p_oSBOApplication.Forms.ActiveForm
                                            oform.Items.Item("5").Specific.String = sBPcode
                                            oform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            oform.Items.Item("btnapprove").Enabled = False

                                        End If
                                    End If
                                End If

                            Catch ex As Exception
                                p_oSBOApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                BubbleEvent = False
                                Exit Try
                            End Try
                            Exit Sub
                        End If

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.ItemUID = "1" Then
                            Dim oform As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount(pVal.FormTypeEx, pVal.FormTypeCount)
                            Dim oform_UDF As SAPbouiCOM.Form = Nothing
                            Try
                                If oform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                    Dim sUDF As String = String.Empty
                                    p_FormTypecount = pVal.FormTypeCount

                                    oform_UDF = p_oSBOApplication.Forms.GetFormByTypeAndCount(-134, pVal.FormTypeCount)
                                    sUDF = oform_UDF.Items.Item("U_AB_STATUS").Specific.value.ToString.Trim().ToUpper
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("User Authorization " & p_oCompDef.sAuthorization, sFuncName)
                                    Select Case p_oCompDef.sAuthorization
                                        Case "CREATE AND UPDATE"
                                            If sUDF = "PENDING" Then
                                                oform.Items.Item("10002045").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            ElseIf sUDF = "APPROVED" Then
                                                p_oSBOApplication.SetStatusBarMessage("Kindly change the status to ""PENDING"" .....!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                BubbleEvent = False
                                                Exit Sub
                                            Else
                                                p_oSBOApplication.SetStatusBarMessage("Invalid Status change it to ""PENDING"" .....!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                                BubbleEvent = False
                                                Exit Sub
                                            End If
                                        Case Else
                                            p_oSBOApplication.SetStatusBarMessage("Kindly define the role(APPROVE/CREATE AND UPDATE) in the user setup .....!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                            BubbleEvent = False
                                            Exit Sub
                                    End Select
                                ElseIf oform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                    p_BPTypecount = pVal.FormTypeCount
                                End If
                            Catch ex As Exception
                                p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Try
                            End Try
                            Exit Sub
                        End If


                End Select
            End If

            'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch exc As Exception
            BubbleEvent = False
            sErrDesc = exc.Message
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            p_oSBOApplication.StatusBar.SetText(sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            WriteToLogFile(Err.Description, sFuncName)
            ShowErr(sErrDesc)
        End Try

    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent

        Try
            If pVal.BeforeAction = False Then
                Select Case pVal.MenuUID

                    Case "1282"

                        Dim oform As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                        Dim oform_UDF As SAPbouiCOM.Form = Nothing

                        If oform.TypeEx = "134" Then
                            Try
                                oform_UDF = p_oSBOApplication.Forms.GetFormByTypeAndCount(-134, oform.TypeCount)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Menu Event " & p_oCompDef.sAuthorization, sFuncName)
                                If p_oCompDef.sAuthorization.Trim() = "CREATE AND UPDATE" Then
                                    oform_UDF.Items.Item("U_AB_STATUS").Specific.select("PENDING")
                                    oform.Items.Item("btnapprove").Enabled = False
                                End If
                            Catch ex As Exception
                                p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                Call WriteToLogFile(sErrDesc, sFuncName)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                                Exit Sub
                            End Try
                        End If

                        If oform.TypeEx = "142" Or oform.TypeEx = "1470000200" Then
                            Try
                                oform.Items.Item("txtnote").Enabled = False
                            Catch ex As Exception
                                p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                Call WriteToLogFile(sErrDesc, sFuncName)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                                Exit Sub
                            End Try
                        End If

                    Case "1289", "1290", "1291", "1288"
                        Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                        Try

                            Dim oform_UDF As SAPbouiCOM.Form = Nothing
                            If oForm.TypeEx = "134" Or oForm.TypeEx = "-134" Then
                                oform_UDF = p_oSBOApplication.Forms.GetFormByTypeAndCount(-134, oForm.TypeCount)
                                oForm = p_oSBOApplication.Forms.GetFormByTypeAndCount(134, oForm.TypeCount)
                                If p_oCompDef.sAuthorization = "APPROVE" And oform_UDF.Items.Item("U_AB_STATUS").Specific.value.ToString.Trim() <> "APPROVED" Then
                                    oForm.Items.Item("btnapprove").Enabled = True
                                Else
                                    oForm.Items.Item("btnapprove").Enabled = False
                                End If

                            ElseIf oForm.TypeEx = "142" Or oForm.TypeEx = "1470000200" Then
                                Try
                                    oForm.Items.Item("txtnote").Enabled = False
                                Catch ex As Exception
                                    p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    Call WriteToLogFile(sErrDesc, sFuncName)
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                                    Exit Sub
                                End Try
                            End If
                        Catch ex As Exception
                            p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                            Exit Sub
                        End Try

                    Case "1281"
                        Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                        If oForm.TypeEx = "142" Or oForm.TypeEx = "1470000200" Then
                            Try
                                oForm.Items.Item("txtnote").Enabled = False
                            Catch ex As Exception
                                p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                Call WriteToLogFile(sErrDesc, sFuncName)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                                Exit Sub
                            End Try
                        End If
                End Select
            End If


        Catch ex As Exception
            p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        End Try
    End Sub

End Class

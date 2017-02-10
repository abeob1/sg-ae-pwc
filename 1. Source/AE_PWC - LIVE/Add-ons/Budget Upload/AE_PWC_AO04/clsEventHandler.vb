Option Explicit On
Imports System.Windows.Forms

Namespace AE_PWC_AO04
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

        Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
            ' **********************************************************************************
            '   Function   :    SBO_Application_MenuEvent()
            '   Purpose    :    This function will be handling the SAP Menu Event
            '               
            '   Parameters :    ByRef pVal As SAPbouiCOM.MenuEvent
            '                       pVal = set the SAP UI MenuEvent Object
            '                   ByRef BubbleEvent As Boolean
            '                       BubbleEvent = set the True/False        
            ' **********************************************************************************
            ' Dim oForm As SAPbouiCOM.Form = Nothing
            Dim sErrDesc As String = String.Empty
            Dim sFuncName As String = String.Empty
            Dim oForm As SAPbouiCOM.Form = Nothing
            Try
                sFuncName = "SBO_Application_MenuEvent()"
                'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

                If Not p_oDICompany.Connected Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectDICompSSO()", sFuncName)
                    If ConnectDICompSSO(p_oDICompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                End If

                If pVal.BeforeAction = False Then
                    Select Case pVal.MenuUID

                        Case "BUPL"
                            Try
                                LoadFromXML("BudgetUpload.srf", SBO_Application)
                                oForm = p_oSBOApplication.Forms.Item("BUP")
                                oForm.Visible = True

                                Exit Try
                            Catch ex As Exception
                                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                            End Try
                            Exit Sub
                    End Select
                End If

                ' If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Catch exc As Exception
                BubbleEvent = False
                ShowErr(exc.Message)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                WriteToLogFile(Err.Description, sFuncName)
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
            Dim oDSTemplateInformation As DataSet = Nothing
            Dim oDTGatheredInformation As New DataTable

            Try
                sFuncName = "SBO_Application_ItemEvent()"
                ' If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

                If Not IsNothing(p_oDICompany) Then
                    If Not p_oDICompany.Connected Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectDICompSSO()", sFuncName)
                        If ConnectDICompSSO(p_oDICompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    End If
                End If

                If pVal.Before_Action = True Then

                    Select Case pVal.FormUID

                        Case "BUP"
                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                                If pVal.ItemUID = "Item_8" And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                    Try
                                        Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling File Open Function", sFuncName)
                                        oForm.Items.Item("4").Specific.string = fillopen()
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With Success File Open Function", sFuncName)
                                        ' oForm.Items.Item("Item_5").Specific.string = p_sSelectedFilepath
                                        Exit Sub

                                    Catch ex As Exception
                                        p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        BubbleEvent = False
                                        Exit Sub
                                    End Try

                                End If

                                If pVal.ItemUID = "btnGntFile" And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                    Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                                    Try
                                        Dim sFilePath As String = String.Empty
                                        Dim SsqlString As String = String.Empty
                                        Dim dSplitAmount As Decimal = 0
                                        Dim oRest As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                        sFilePath = oForm.Items.Item("4").Specific.String
                                        oForm.Items.Item("1000001").Specific.String = String.Empty

                                        If Not String.IsNullOrEmpty(sFilePath) Then
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling GetDataViewFromCSV Function", sFuncName)
                                            WriteIntoEditBox(oForm, "Calling GetDataViewFromCSV() ", sErrDesc)
                                            oDSTemplateInformation = GetDatasetFromExcel(oForm, sFilePath, sErrDesc)

                                            If oDSTemplateInformation Is Nothing Then
                                                WriteIntoEditBox(oForm, "No Datas in the Budget Template file .......", sErrDesc)
                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No Datas in the Budget Template file", sFuncName)
                                                BubbleEvent = False
                                                Exit Sub
                                            End If

                                            Select Case p_sBudgetType
                                                Case "OU"
                                                    WriteIntoEditBox(oForm, "Calling StreamlineBudgetInformation_OU() ", sErrDesc)
                                                    oDTGatheredInformation = StreamlineBudgetInformation_OU(oForm, oDSTemplateInformation, sErrDesc)
                                                    ''Case "BU"
                                                    ''    WriteIntoEditBox(oForm, "Calling StreamlineBudgetInformation_BU() ", sErrDesc)
                                                    ''    oDTGatheredInformation = StreamlineBudgetInformation_BU(oForm, oDSTemplateInformation, sErrDesc)
                                                Case "IF"
                                                    WriteIntoEditBox(oForm, "Calling StreamlineBudgetInformation_PR() ", sErrDesc)
                                                    oDTGatheredInformation = StreamlineBudgetInformation_PR(oForm, oDSTemplateInformation, sErrDesc)
                                            End Select
                                           

                                            If oDTGatheredInformation Is Nothing Or oDTGatheredInformation.Rows.Count = 0 Then
                                                WriteIntoEditBox(oForm, "No Datas in the Budget Template file .......", sErrDesc)
                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No Datas in the Budget Template file", sFuncName)
                                                BubbleEvent = False
                                                Exit Sub
                                            Else
                                                WriteIntoEditBox(oForm, "Inserting Values In the Budget Table ", sErrDesc)
                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting Values In the Budget Table ", sFuncName)
                                                For Each DR As DataRow In oDTGatheredInformation.Rows

                                                    '                                                           Budget Type                                  Budget NAme                                   Budget Period                                    OU Code                                     BU Code                                 Project Code                                  Budget Amount                        
                                                    SsqlString = "PWCL.. [@AE_SP002_InsertintoBudgetTable]'" & DR.Item("BudgetType").ToString.Trim & "', '" & DR.Item("BudgetName").ToString.Trim & "','" & DR.Item("BudgetPeriod").ToString.Trim & "','" & DR.Item("CostCenter").ToString.Trim & "','" & DR.Item("BUCode").ToString.Trim & "','" & DR.Item("ProjectCode").ToString.Trim & "', '" & DR.Item("BudgetAmount").ToString.Trim & "', " & _
                                                         "" & DR.Item("GLAccount").ToString.Trim & "," & CDbl(DR.Item("Month1").ToString.Trim) & "," & CDbl(DR.Item("Month2").ToString.Trim) & "," & CDbl(DR.Item("Month3").ToString.Trim) & "," & CDbl(DR.Item("Month4").ToString.Trim) & " " & _
                                                         "," & CDbl(DR.Item("Month5").ToString.Trim) & "," & CDbl(DR.Item("Month6").ToString.Trim) & "," & CDbl(DR.Item("Month7").ToString.Trim) & "," & CDbl(DR.Item("Month8").ToString.Trim) & "," & CDbl(DR.Item("Month9").ToString.Trim) & "," & CDbl(DR.Item("Month10").ToString.Trim) & " " & _
                                                             "," & CDbl(DR.Item("Month11").ToString.Trim) & "," & CDbl(DR.Item("Month12").ToString.Trim) & ", " & DR.Item("sAmount").ToString.Trim & ""


                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Insert SQL " & SsqlString, sFuncName)
                                                    oRest.DoQuery(SsqlString)
                                                Next
                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Successfully Inserted In the Budget Table", sFuncName)
                                                WriteIntoEditBox(oForm, "Successfully Inserted In the Budget Table ", sErrDesc)
                                            End If
                                        Else
                                            p_oSBOApplication.StatusBar.SetText("Choose the budget template to upload .....!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                            BubbleEvent = False
                                            Exit Sub
                                        End If

                                    Catch ex As Exception
                                        WriteIntoEditBox(oForm, "Error : " & ex.Message, sErrDesc)
                                        p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        BubbleEvent = False
                                        Exit Sub
                                    End Try
                                End If
                            End If
                    End Select
                End If


                'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Catch exc As Exception
                BubbleEvent = False
                sErrDesc = exc.Message
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                WriteToLogFile(Err.Description, sFuncName)
                ShowErr(sErrDesc)
            End Try

        End Sub

        Sub AddMenuItems()
            Dim oMenus As SAPbouiCOM.Menus
            Dim oMenuItem As SAPbouiCOM.MenuItem
            oMenus = SBO_Application.Menus

            Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
            oCreationPackage = (SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams))
            oMenuItem = SBO_Application.Menus.Item("43520") 'Modules

            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
            oCreationPackage.UniqueID = "PWC"
            oCreationPackage.String = "Customization"
            oCreationPackage.Enabled = True
            oCreationPackage.Position = -1

            oCreationPackage.Image = System.Windows.Forms.Application.StartupPath & "\Logo.bmp"
            oMenus = oMenuItem.SubMenus

            Try
                'If the manu already exists this code will fail
                If Not p_oSBOApplication.Menus.Exists("PWC") Then
                    oMenus.AddEx(oCreationPackage)
                End If

            Catch
            End Try


            Try
                'Get the menu collection of the newly added pop-up item
                oMenuItem = SBO_Application.Menus.Item("PWC")
                oMenus = oMenuItem.SubMenus

                'Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oCreationPackage.UniqueID = "BUPL"
                oCreationPackage.String = "Budget Upload"

                If Not p_oSBOApplication.Menus.Exists("BUPL") Then
                    oMenus.AddEx(oCreationPackage)
                End If



            Catch
                'Menu already exists
                SBO_Application.SetStatusBarMessage("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, False)
            End Try
        End Sub

        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub

        Private Sub SBO_Application_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles SBO_Application.FormDataEvent

            Select Case BusinessObjectInfo.FormTypeEx

                Case "804"

                    If BusinessObjectInfo.ActionSuccess = True And BusinessObjectInfo.BeforeAction = False And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD) Then ' Or BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) Then


                        Dim ival As Integer
                        Dim IsError As Boolean
                        Dim iErr As Integer = 0
                        Dim sErr As String = String.Empty

                        Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount(804, p_FrmType)
                        Dim sAcctCode As String = String.Empty
                        sAcctCode = oForm.DataSources.DBDataSources.Item(0).GetValue("AcctCode", "0").ToString.Trim
                        Dim oChartofAccounts As SAPbobsCOM.ChartOfAccounts = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oChartOfAccounts)
                        If oChartofAccounts.GetByKey(sAcctCode) Then
                            oChartofAccounts.FrozenFor = SAPbobsCOM.BoYesNoEnum.tYES
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Set InActive for Account Code " & sAcctCode, sFuncName)
                            ival = oChartofAccounts.Update()
                            If ival <> 0 Then
                                IsError = True
                                p_oDICompany.GetLastError(iErr, sErr)
                                p_oSBOApplication.StatusBar.SetText(sErr, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
                                Exit Sub
                            End If
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS " & sAcctCode, sFuncName)
                        End If

                    End If


            End Select

        End Sub

        Private Sub SBO_Application_LayoutKeyEvent(ByRef eventInfo As SAPbouiCOM.LayoutKeyInfo, ByRef BubbleEvent As Boolean) Handles SBO_Application.LayoutKeyEvent

        End Sub
    End Class
End Namespace



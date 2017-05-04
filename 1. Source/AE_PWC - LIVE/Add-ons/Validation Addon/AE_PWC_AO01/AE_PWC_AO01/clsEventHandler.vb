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

    ' ''Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
    ' ''    ' **********************************************************************************
    ' ''    '   Function   :    SBO_Application_MenuEvent()
    ' ''    '   Purpose    :    This function will be handling the SAP Menu Event
    ' ''    '               
    ' ''    '   Parameters :    ByRef pVal As SAPbouiCOM.MenuEvent
    ' ''    '                       pVal = set the SAP UI MenuEvent Object
    ' ''    '                   ByRef BubbleEvent As Boolean
    ' ''    '                       BubbleEvent = set the True/False        
    ' ''    ' **********************************************************************************
    ' ''    Dim oForm As SAPbouiCOM.Form = Nothing
    ' ''    Dim sErrDesc As String = String.Empty
    ' ''    Dim sFuncName As String = String.Empty

    ' ''    Try
    ' ''        sFuncName = "SBO_Application_MenuEvent()"
    ' ''        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

    ' ''        If Not p_oDICompany.Connected Then
    ' ''            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectDICompSSO()", sFuncName)
    ' ''            If ConnectDICompSSO(p_oDICompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
    ' ''        End If

    ' ''        If pVal.BeforeAction = False Then
    ' ''            Select Case pVal.MenuUID
    ' ''                Case "IS"
    ' ''                    Try
    ' ''                        LoadFromXML("ImportStatistcs.srf", SBO_Application)
    ' ''                        oForm = SBO_Application.Forms.Item("ImpStat")

    ' ''                        oForm.Visible = True
    ' ''                        Exit Try

    ' ''                    Catch ex As Exception
    ' ''                        SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    ' ''                        BubbleEvent = False
    ' ''                    End Try
    ' ''                    Exit Sub

    ' ''            End Select

    ' ''            ''Else
    ' ''            ''    Select Case pVal.MenuUID

    ' ''            ''        Case "1286", "1284" 'Close & Cancel Document leve
    ' ''            ''            oForm = p_oSBOApplication.Forms.ActiveForm
    ' ''            ''            If oForm.Title = "Purchase Order [Approved]" Then
    ' ''            ''                Try
    ' ''            ''                    Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("38").Specific
    ' ''            ''                    For imjs As Integer = 1 To oMatrix.VisualRowCount
    ' ''            ''                        If Not String.IsNullOrEmpty(oMatrix.Columns.Item("1").Cells.Item(imjs).Specific.String) Then
    ' ''            ''                            If CDbl(oMatrix.Columns.Item("11").Cells.Item(imjs).Specific.String) = CDbl(oMatrix.Columns.Item("32").Cells.Item(imjs).Specific.String) Then
    ' ''            ''                                p_oSBOApplication.StatusBar.SetText("Can`t close / cancel this document, some lines don`t have the traget document ............ ! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    ' ''            ''                                BubbleEvent = False
    ' ''            ''                                Exit Sub
    ' ''            ''                            End If
    ' ''            ''                        End If
    ' ''            ''                    Next
    ' ''            ''                Catch ex As Exception
    ' ''            ''                    SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    ' ''            ''                    BubbleEvent = False
    ' ''            ''                End Try
    ' ''            ''                Exit Sub
    ' ''            ''            End If

    ' ''            ''        Case "1299", "1293" ' Delete & close Row level
    ' ''            ''            oForm = p_oSBOApplication.Forms.ActiveForm
    ' ''            ''            If oForm.Title = "Purchase Order [Approved]" Then
    ' ''            ''                Try
    ' ''            ''                    Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("38").Specific
    ' ''            ''                    Dim bMAtrixSelection As Boolean = False
    ' ''            ''                    For imjs As Integer = 1 To oMatrix.VisualRowCount
    ' ''            ''                        If oMatrix.IsRowSelected(imjs) Then
    ' ''            ''                            bMAtrixSelection = True
    ' ''            ''                            If CDbl(oMatrix.Columns.Item("11").Cells.Item(imjs).Specific.String) = CDbl(oMatrix.Columns.Item("32").Cells.Item(imjs).Specific.String) Then
    ' ''            ''                                p_oSBOApplication.StatusBar.SetText("Can`t close / cancel this document, some lines don`t have the traget document ............ ! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    ' ''            ''                                BubbleEvent = False
    ' ''            ''                                Exit Sub
    ' ''            ''                            End If
    ' ''            ''                        End If
    ' ''            ''                    Next

    ' ''            ''                    If bMAtrixSelection = False Then
    ' ''            ''                        p_oSBOApplication.StatusBar.SetText("Kindly select the row before click 'Close Row' ............ ! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    ' ''            ''                        BubbleEvent = False
    ' ''            ''                        Exit Sub
    ' ''            ''                    End If

    ' ''            ''                Catch ex As Exception
    ' ''            ''                    SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    ' ''            ''                    BubbleEvent = False
    ' ''            ''                End Try
    ' ''            ''                Exit Sub
    ' ''            ''            End If

    ' ''            ''    End Select
    ' ''        End If

    ' ''        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
    ' ''    Catch exc As Exception
    ' ''        BubbleEvent = False
    ' ''        ShowErr(exc.Message)
    ' ''        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
    ' ''        WriteToLogFile(Err.Description, sFuncName)
    ' ''    End Try
    ' ''End Sub

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

                ''Select Case pVal.FormUID
                ''    Case "ImpStat"
                ''        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                ''            If pVal.ItemUID = "btnBrowse" Then
                ''                Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                ''                sFuncName = "'Browse' Button Click - ID 'btnBrowse'"
                ''                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling File Open Function", sFuncName)

                ''                oForm.Items.Item("txtPath").Specific.string = fillopen()

                ''                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With Success File Open Function", sFuncName)
                ''                ' oForm.Items.Item("Item_5").Specific.string = p_sSelectedFilepath
                ''                Exit Sub
                ''            End If

                ''            If pVal.ItemUID = "btnImport" Then
                ''                Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                ''                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ImportStatistics()", sFuncName)

                ''                p_oSBOApplication.StatusBar.SetText("Please Wait Importing the Data...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                ''                If ImportStatistics(oForm, sErrDesc, BubbleEvent) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                ''                ' oForm.Items.Item("txtPath").Specific.string = String.Empty
                ''                p_oSBOApplication.StatusBar.SetText("Data Imported Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                ''            End If

                ''        End If
                ''End Select

                Select Case pVal.FormTypeEx



                    ''Case "10001"

                    ''    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                    ''        Dim oform As SAPbouiCOM.Form = Nothing
                    ''        Dim oform_UDF As SAPbouiCOM.Form = Nothing
                    ''        Try
                    ''            If p_BPTypecount > 0 Then
                    ''                oform = p_oSBOApplication.Forms.GetFormByTypeAndCount(134, p_FormTypecount)
                    ''                oform_UDF = p_oSBOApplication.Forms.GetFormByTypeAndCount(-134, p_FormTypecount)
                    ''                If p_oCompDef.sAuthorization = "APPROVE" And oform_UDF.Items.Item("U_AB_STATUS").Specific.value.ToString.Trim() <> "APPROVED" Then
                    ''                    oform.Items.Item("btnapprove").Enabled = True
                    ''                Else
                    ''                    oform.Items.Item("btnapprove").Enabled = False
                    ''                End If
                    ''                p_BPTypecount = 0
                    ''            End If

                    ''        Catch ex As Exception
                    ''            p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    ''            Call WriteToLogFile(sErrDesc, sFuncName)
                    ''            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                    ''            Exit Sub
                    ''        End Try
                    ''    End If

                    Case "10021", "1470010112"

                        ''If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                        ''    Dim oform As SAPbouiCOM.Form = Nothing
                        ''    Dim iformtype As Integer = 0
                        ''    Try
                        ''        If p_FormTypecount > 0 Then
                        ''            If pVal.FormTypeEx = "10021" Then
                        ''                iformtype = 142
                        ''            Else
                        ''                iformtype = 1470000200
                        ''            End If
                        ''            oform = p_oSBOApplication.Forms.GetFormByTypeAndCount(iformtype, p_FormTypecount)
                        ''            oform.Items.Item("txtnote").Enabled = False
                        ''            p_FormTypecount = 0
                        ''        End If

                        ''    Catch ex As Exception
                        ''        p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        ''        Call WriteToLogFile(sErrDesc, sFuncName)
                        ''        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                        ''        Exit Sub
                        ''    End Try
                        ''End If

                        ''Case "-134"
                        ''    If pVal.ItemChanged = True Then

                        ''        Try
                        ''            Dim oform As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount(134, pVal.FormTypeCount)
                        ''            Dim oform_UDF As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount(-134, pVal.FormTypeCount)
                        ''            Dim sStatus As String = String.Empty
                        ''            sStatus = oform_UDF.Items.Item("U_AB_STATUS").Specific.value.ToString.Trim()

                        ''            If p_oCompDef.sAuthorization = "APPROVE" And sStatus <> "APPROVED" Then
                        ''                oform.Items.Item("btnapprove").Enabled = True
                        ''            Else
                        ''                oform.Items.Item("btnapprove").Enabled = False
                        ''            End If
                        ''        Catch ex As Exception
                        ''            p_oSBOApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        ''            BubbleEvent = False
                        ''            Exit Try
                        ''        End Try
                        ''    End If

                        ''Case "134"
                        ''    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD Then
                        ''        Dim oform As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount(pVal.FormTypeEx, pVal.FormTypeCount)
                        ''        Dim oItem As SAPbouiCOM.Item = Nothing
                        ''        Dim oRItem As SAPbouiCOM.Item = Nothing

                        ''        oItem = oform.Items.Add("btnapprove", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
                        ''        oRItem = oform.Items.Item("2")
                        ''        oItem.Height = oRItem.Height
                        ''        oItem.Width = oRItem.Width
                        ''        oItem.Top = oRItem.Top
                        ''        oItem.Left = oRItem.Left + oRItem.Width + 5
                        ''        oItem.Visible = True
                        ''        oform.Items.Item("btnapprove").Specific.caption = "Approve"
                        ''        If p_oCompDef.sAuthorization = "APPROVE" Then
                        ''            oform.Items.Item("btnapprove").Enabled = True
                        ''        Else
                        ''            oform.Items.Item("btnapprove").Enabled = False

                        ''        End If
                        ''    End If

                        ''    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.ItemUID = "1" Then
                        ''        Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount(134, pVal.FormTypeCount)
                        ''        Dim oform_UDF As SAPbouiCOM.Form = Nothing
                        ''        Try

                        ''            oform_UDF = p_oSBOApplication.Forms.GetFormByTypeAndCount(-134, oForm.TypeCount)
                        ''            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        ''                If p_oCompDef.sAuthorization = "APPROVE" And oform_UDF.Items.Item("U_AB_STATUS").Specific.value.ToString.Trim() <> "APPROVED" Then
                        ''                    oForm.Items.Item("btnapprove").Enabled = True
                        ''                Else
                        ''                    oForm.Items.Item("btnapprove").Enabled = False
                        ''                End If
                        ''            End If
                        ''        Catch ex As Exception
                        ''            p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        ''            Call WriteToLogFile(sErrDesc, sFuncName)
                        ''            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                        ''            BubbleEvent = False
                        ''            Exit Sub
                        ''        End Try


                        ''    End If
                    Case "142"


                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DRAW Then
                            Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount(142, pVal.FormTypeCount)
                            Dim oComboSeries As SAPbouiCOM.ComboBox
                            Dim oNewForm As SAPbouiCOM.Form = Nothing

                            Try
                                sFuncName = "Validating the Pre Approval"
                                oComboSeries = oForm.Items.Item("88").Specific
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)
                                oNewForm = p_oSBOApplication.Forms.GetFormByTypeAndCount("-142", pVal.FormTypeCount)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Selected Series  " & oComboSeries.Selected.Description.ToUpper().Trim(), sFuncName)
                                If oComboSeries.Selected.Description.ToUpper().Trim() = "INTERNAL" Or oComboSeries.Selected.Description.ToUpper().Trim() = "ANNUAL" Then
                                    oNewForm.Items.Item("U_AB_PREAPPROVED").Enabled = True
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Enabling the Field", sFuncName)
                                Else
                                    oNewForm.Items.Item("U_AB_PREAPPROVED").Enabled = False
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disabling the Field", sFuncName)
                                End If
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
                            Catch ex As Exception
                                p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                Call WriteToLogFile(sErrDesc, sFuncName)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                                BubbleEvent = False
                                Exit Sub
                            End Try
                        End If

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT And pVal.ItemUID = "88" Then
                            Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount(142, pVal.FormTypeCount)
                            Dim oComboSeries As SAPbouiCOM.ComboBox
                            Dim oNewForm As SAPbouiCOM.Form = Nothing
                            Try
                                sFuncName = "Validating the Pre Approval"
                                oComboSeries = oForm.Items.Item("88").Specific
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)
                                oNewForm = p_oSBOApplication.Forms.GetFormByTypeAndCount("-142", pVal.FormTypeCount)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Selected Series  " & oComboSeries.Selected.Description.ToUpper().Trim(), sFuncName)
                                If oComboSeries.Selected.Description.ToUpper().Trim() = "INTERNAL" Or oComboSeries.Selected.Description.ToUpper().Trim() = "ANNUAL" Then
                                    oNewForm.Items.Item("U_AB_PREAPPROVED").Enabled = True
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Enabling the Field", sFuncName)
                                Else
                                    oNewForm.Items.Item("U_AB_PREAPPROVED").Enabled = False
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disabling the Field", sFuncName)
                                End If
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
                            Catch ex As Exception
                                p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                Call WriteToLogFile(sErrDesc, sFuncName)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                                BubbleEvent = False
                                Exit Sub
                            End Try
                        End If

                    Case "-142"
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT And pVal.ItemUID = "U_AB_PREAPPROVED" Then
                            Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount(-142, pVal.FormTypeCount)
                            Dim oComboPreApproved As SAPbouiCOM.ComboBox
                            Dim oComboSeries As SAPbouiCOM.ComboBox
                            Dim oNewForm As SAPbouiCOM.Form = Nothing
                            oNewForm = p_oSBOApplication.Forms.GetFormByTypeAndCount("142", pVal.FormTypeCount)

                            If p_PREAPPROVED = True Then Exit Sub

                            Dim oMatrix As SAPbouiCOM.Matrix = oNewForm.Items.Item("38").Specific
                            Try
                                oComboSeries = oNewForm.Items.Item("88").Specific
                                oComboPreApproved = oForm.Items.Item("U_AB_PREAPPROVED").Specific
                                If oComboSeries.Selected.Description.ToUpper().Trim() = "INTERNAL" Or oComboSeries.Selected.Description.ToUpper().Trim() = "ANNUAL" Then
                                    If oComboPreApproved.Selected.Value = "Y" And String.IsNullOrEmpty(oMatrix.Columns.Item("44").Cells.Item(1).Specific.String) Then
                                        SBO_Application.MessageBox(" Reminder :" & vbCrLf & "You have opted for Pre-approved, please attach evidence of approval", 1, "Ok")
                                        BubbleEvent = False
                                        Exit Sub
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

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                            If pVal.ItemUID = "3" Then
                                Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
                                Try
                                    Dim oMAtrix As SAPbouiCOM.Matrix = Nothing
                                    oMAtrix = oForm.Items.Item("3").Specific
                                    ''   p_sStatus = oMAtrix.Columns.Item("30").Cells.Item(pVal.Row).Specific.String
                                    p_sAppStatus = oMAtrix.Columns.Item("30").Cells.Item(pVal.Row).Specific.String
                                Catch ex As Exception
                                    p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    Call WriteToLogFile(sErrDesc, sFuncName)
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                                    BubbleEvent = False
                                    Exit Sub
                                End Try
                            End If
                        End If

                    Case "392", "393"

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT And pVal.ItemUID = "137" Then
                            Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetForm(pVal.FormType, pVal.FormTypeCount)

                            Dim oCombo As SAPbouiCOM.ComboBox
                            Dim oCheckBox As SAPbouiCOM.CheckBox

                            oCombo = oForm.Items.Item("137").Specific
                            oCheckBox = oForm.Items.Item("99").Specific

                            P_JEReversetmp = True
                            If oCombo.Selected.Description = "GJR" Then
                                P_JEReverse = True
                                oCheckBox.Checked = True
                                ReverseDate_Validation(oForm, sErrDesc)
                                oForm.Items.Item("98").Specific.string = dtReverseDate
                            Else
                                If oCheckBox.Checked = True Then
                                    oForm.Items.Item("98").Specific.string = ""
                                    oCheckBox.Checked = False
                                Else
                                    P_JEReversetmp = False
                                End If
                            End If
                        End If

                        ''Case "142", "1470000200"

                        ''    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD Then
                        ''        Dim oform As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount(pVal.FormTypeEx, pVal.FormTypeCount)
                        ''        Dim oItem As SAPbouiCOM.Item = Nothing
                        ''        Dim oRItem As SAPbouiCOM.Item = Nothing
                        ''        Dim sString As String = String.Empty

                        ''        Try
                        ''            sString = "1)        Terms & Conditions (T&C): In most instances, PwC's T&C should be used. Please consult OGC if we sign any suppliers’ agreement. For details, please refer to Finance Policies and Procedures (Procurement section). "
                        ''            sString += vbCrLf
                        ''            sString += "2)        Procurement of IT systems and applications: Prior to making any purchase of IT systems/ applications/ services, please consult GTS managers for Infrastructure and Personal Computing. For details, please refer to Finance Policies and Procedures (Procurement section). "

                        ''            oItem = oform.Items.Add("lbl", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                        ''            oRItem = oform.Items.Item("230")
                        ''            oItem.Height = oRItem.Height
                        ''            oItem.Width = oRItem.Width
                        ''            oItem.Top = oRItem.Top + oRItem.Height + 3
                        ''            oItem.Left = oRItem.Left
                        ''            oItem.Visible = True
                        ''            oform.Items.Item("lbl").Specific.caption = "Note  "


                        ''            oItem = oform.Items.Add("txtnote", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)
                        ''            oRItem = oform.Items.Item("222")
                        ''            oItem.Height = 50
                        ''            oItem.Width = 650
                        ''            oItem.Top = oRItem.Top + oRItem.Height + 3
                        ''            oItem.Left = oRItem.Left
                        ''            oItem.Visible = True
                        ''            oItem.Enabled = False
                        ''            oform.Items.Item("txtnote").Specific.String = sString
                        ''            oform.Items.Item("lbl").LinkTo = "txtnote"

                        ''        Catch ex As Exception
                        ''            p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        ''            BubbleEvent = False
                        ''            Exit Sub
                        ''        End Try
                        ''    End If

                        ''    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_RESIZE Then
                        ''        Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
                        ''        Try
                        ''            oForm.Freeze(True)
                        ''            If oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Minimized Then
                        ''                oForm.Items.Item("txtnote").Width = 150
                        ''                oForm.Items.Item("txtnote").Top = oForm.Items.Item("222").Top + oForm.Items.Item("222").Height + 3
                        ''                oForm.Items.Item("lbl").LinkTo = "txtnote"
                        ''            ElseIf oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized Then
                        ''                oForm.Items.Item("txtnote").Width = 650
                        ''                oForm.Items.Item("txtnote").Top = oForm.Items.Item("222").Top + oForm.Items.Item("222").Height + 3
                        ''                oForm.Items.Item("lbl").LinkTo = "txtnote"
                        ''            ElseIf oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Restore Then
                        ''                If oForm.Items.Item("txtnote").Width = 150 Then
                        ''                    oForm.Items.Item("txtnote").Width = 650
                        ''                    oForm.Items.Item("txtnote").Top = oForm.Items.Item("222").Top + oForm.Items.Item("222").Height + 3
                        ''                    oForm.Items.Item("lbl").LinkTo = "txtnote"
                        ''                Else
                        ''                    oForm.Items.Item("txtnote").Width = 150
                        ''                    oForm.Items.Item("txtnote").Top = oForm.Items.Item("222").Top + oForm.Items.Item("222").Height + 3
                        ''                    oForm.Items.Item("lbl").LinkTo = "txtnote"
                        ''                End If
                        ''            End If
                        ''            oForm.Freeze(False)
                        ''        Catch ex As Exception
                        ''            oForm.Freeze(False)
                        ''            Call WriteToLogFile(sErrDesc, sFuncName)
                        ''            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                        ''            p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        ''            BubbleEvent = False
                        ''            Exit Sub
                        ''        End Try
                        ''    End If

                        ''    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.ItemUID = "1" Then
                        ''        Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
                        ''        Try
                        ''            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        ''                oForm.Items.Item("txtnote").Enabled = False
                        ''            End If
                        ''        Catch ex As Exception
                        ''            p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        ''            Call WriteToLogFile(sErrDesc, sFuncName)
                        ''            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                        ''            BubbleEvent = False
                        ''            Exit Sub
                        ''        End Try


                        ''    End If

                End Select
            Else   ' Before action True

                Select Case pVal.FormTypeEx

                    '----- Commented BP Approval as per the NIK advise - 29-01-2015
                    ''Case "10001"

                    ''    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then

                    ''        Dim oForm As SAPbouiCOM.Form = Nothing
                    ''        Dim oForm_L As SAPbouiCOM.Form = Nothing
                    ''        Dim oform_UDF As SAPbouiCOM.Form = Nothing
                    ''        Try
                    ''            oForm_L = p_oSBOApplication.Forms.ActiveForm
                    ''            If oForm_L.Title.ToUpper = "LIST OF BUSINESS PARTNERS" And p_BPTypecount > 0 Then
                    ''                oForm = p_oSBOApplication.Forms.GetFormByTypeAndCount(134, p_BPTypecount)
                    ''                oform_UDF = p_oSBOApplication.Forms.GetFormByTypeAndCount(-134, p_BPTypecount)
                    ''                If p_oCompDef.sAuthorization = "APPROVE" And oform_UDF.Items.Item("U_AB_STATUS").Specific.value.ToString.Trim() <> "APPROVED" Then
                    ''                    oForm.Items.Item("btnapprove").Enabled = True
                    ''                Else
                    ''                    oForm.Items.Item("btnapprove").Enabled = False
                    ''                End If
                    ''                p_BPTypecount = 0
                    ''            End If

                    ''        Catch ex As Exception
                    ''            p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    ''            BubbleEvent = False
                    ''            Exit Sub
                    ''        End Try
                    ''    End If

                  

                    Case "134" ' BP Approval

                        ''If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.ItemUID = "btnapprove" Then
                        ''    Dim oform As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount(pVal.FormTypeEx, pVal.FormTypeCount)
                        ''    '' Dim oform_UDF As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount(-134, pVal.FormTypeCount)
                        ''    ''oform.Items.Item("10002044").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        ''    Dim sBPcode As String = String.Empty
                        ''    Dim oBP As SAPbobsCOM.BusinessPartners = Nothing
                        ''    Dim irlt As Integer = 0
                        ''    Try
                        ''        If oform.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE And p_oCompDef.sAuthorization = "APPROVE" Then
                        ''            oBP = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
                        ''            sBPcode = oform.Items.Item("5").Specific.String
                        ''            If oBP.GetByKey(sBPcode) Then
                        ''                oBP.Frozen = SAPbobsCOM.BoYesNoEnum.tNO
                        ''                oBP.Valid = SAPbobsCOM.BoYesNoEnum.tYES
                        ''                oBP.UserFields.Fields.Item("U_AB_STATUS").Value = "APPROVED"
                        ''                irlt = oBP.Update()
                        ''                If irlt <> 0 Then
                        ''                    p_oSBOApplication.SetStatusBarMessage(p_oDICompany.GetLastErrorCode & " - " & p_oDICompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        ''                    BubbleEvent = False
                        ''                    Exit Try
                        ''                Else
                        ''                    oform.Close()
                        ''                    p_oSBOApplication.ActivateMenuItem("2561")
                        ''                    oform = p_oSBOApplication.Forms.ActiveForm
                        ''                    oform.Items.Item("5").Specific.String = sBPcode
                        ''                    oform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        ''                    oform.Items.Item("btnapprove").Enabled = False

                        ''                End If
                        ''            End If
                        ''        End If

                        ''    Catch ex As Exception
                        ''        p_oSBOApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        ''        BubbleEvent = False
                        ''        Exit Try
                        ''    End Try
                        ''    Exit Sub
                        ''End If

                        ''If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.ItemUID = "1" Then
                        ''    Dim oform As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount(pVal.FormTypeEx, pVal.FormTypeCount)
                        ''    Dim oform_UDF As SAPbouiCOM.Form = Nothing
                        ''    Try
                        ''        If oform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        ''            Dim sUDF As String = String.Empty
                        ''            p_FormTypecount = pVal.FormTypeCount

                        ''            oform_UDF = p_oSBOApplication.Forms.GetFormByTypeAndCount(-134, pVal.FormTypeCount)
                        ''            sUDF = oform_UDF.Items.Item("U_AB_STATUS").Specific.value.ToString.Trim().ToUpper
                        ''            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("User Authorization " & p_oCompDef.sAuthorization, sFuncName)
                        ''            Select Case p_oCompDef.sAuthorization
                        ''                Case "CREATE AND UPDATE"
                        ''                    If sUDF = "PENDING" Then
                        ''                        oform.Items.Item("10002045").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        ''                    ElseIf sUDF = "APPROVED" Then
                        ''                        p_oSBOApplication.SetStatusBarMessage("Kindly change the status to ""PENDING"" .....!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                        ''                        BubbleEvent = False
                        ''                        Exit Sub
                        ''                    Else
                        ''                        p_oSBOApplication.SetStatusBarMessage("Invalid Status .....!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                        ''                        BubbleEvent = False
                        ''                        Exit Sub
                        ''                    End If
                        ''                Case Else
                        ''                    p_oSBOApplication.SetStatusBarMessage("Kindly define the role(APPROVE/CREATE AND UPDATE) in the user setup .....!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                        ''                    BubbleEvent = False
                        ''                    Exit Sub
                        ''            End Select
                        ''        ElseIf oform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                        ''            p_BPTypecount = pVal.FormTypeCount
                        ''        End If
                        ''    Catch ex As Exception
                        ''        p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        ''        BubbleEvent = False
                        ''        Exit Try
                        ''    End Try
                        ''    Exit Sub
                        ''End If

                    Case "392", "393"
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                            Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
                            If pVal.ItemUID = "1" And oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("76").Specific
                                Try
                                    For imjs As Integer = 1 To oMatrix.VisualRowCount
                                        If Not String.IsNullOrEmpty(oMatrix.Columns.Item("1").Cells.Item(imjs).Specific.String) Then

                                            If String.IsNullOrEmpty(oMatrix.Columns.Item("2006").Cells.Item(imjs).Specific.String) Then
                                                p_oSBOApplication.StatusBar.SetText("Line of service should not be Empty .......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                                BubbleEvent = False
                                                Exit Sub
                                            End If

                                            If String.IsNullOrEmpty(oMatrix.Columns.Item("2001").Cells.Item(imjs).Specific.String) Then
                                                p_oSBOApplication.StatusBar.SetText("Business Unit should not be Empty ........... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                                BubbleEvent = False
                                                Exit Sub
                                            End If

                                            If String.IsNullOrEmpty(oMatrix.Columns.Item("2003").Cells.Item(imjs).Specific.String) Then
                                                p_oSBOApplication.StatusBar.SetText("Operating Unit should not be Empty ........... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                                BubbleEvent = False
                                                Exit Sub
                                            End If
                                        End If

                                    Next imjs
                                Catch ex As Exception
                                    p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                End Try
                                Exit Sub

                            End If
                        End If

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE Then
                            P_JEReversetmp = False
                        End If

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK And pVal.ItemUID = "99" Then
                            Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetForm(pVal.FormType, 0)

                            If P_JEReverse = False Then
                                Dim oCombo As SAPbouiCOM.ComboBox
                                Dim oCheck As SAPbouiCOM.CheckBox

                                oCombo = oForm.Items.Item("137").Specific
                                oCheck = oForm.Items.Item("99").Specific

                                If oCombo.Selected.Description <> "GJR" And P_JEReversetmp = False Then
                                    p_oSBOApplication.StatusBar.SetText("User should not select Reverse", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    BubbleEvent = False
                                    Exit Sub
                                Else
                                    P_JEReversetmp = False
                                End If
                            Else
                                P_JEReverse = False
                                P_JEReversetmp = False
                            End If
                        End If



                    Case "1470000200"
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                            p_FormTypecount = pVal.FormTypeCount
                            Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount(1470000200, pVal.FormTypeCount)
                            Dim oForm_UDF As SAPbouiCOM.Form = Nothing
                            Dim Sql As String = String.Empty
                            If pVal.ItemUID = "1" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                Try
                                    Dim orset As SAPbobsCOM.Recordset = Nothing
                                    Dim sSQL As String = String.Empty
                                    Dim dAmount As Double = 0.0
                                    Dim sPurchasingDepartment As String = String.Empty
                                    p_FormTypecount = pVal.FormTypeCount
                                    oForm_UDF = p_oSBOApplication.Forms.GetFormByTypeAndCount(-1470000200, pVal.FormTypeCount)
                                    orset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                    If String.IsNullOrEmpty(oForm.Items.Item("U_AB_PURCHASEDEPT").Specific.string) Then
                                        p_oSBOApplication.StatusBar.SetText("Approving Unit should not be Empty .............!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        BubbleEvent = False
                                        Exit Try
                                    End If

                                    sPurchasingDepartment = oForm.Items.Item("U_AB_PURCHASEDEPT").Specific.string
                                    If oForm.Title = "Purchase Request - Draft [Approved]" Then
                                        If String.IsNullOrEmpty(oForm.Items.Item("U_AB_POCREATOR").Specific.string) Then
                                            p_oSBOApplication.StatusBar.SetText("PO Creator should not be Empty .............!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                            BubbleEvent = False
                                            Exit Try
                                        Else
                                            sSQL = "SELECT isnull(T0.[Email],'') [Email] FROM OSLP T0 WHERE T0.[SlpName]  = '" & oForm.Items.Item("U_AB_POCREATOR").Specific.string & "'"
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting Email for Purchasing Unit " & sSQL, sFuncName)
                                            orset.DoQuery(sSQL)
                                            If String.IsNullOrEmpty(orset.Fields.Item("Email").Value) Then
                                                p_oSBOApplication.StatusBar.SetText("Email should not blank for this purchasing unit " & oForm.Items.Item("U_AB_POCREATOR").Specific.string, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                                BubbleEvent = False
                                                Exit Try
                                            End If
                                        End If
                                    Else
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling GetPRMaxAmount() " & dAmount, sFuncName)
                                        dAmount = GetPRMaxAmount(oForm, sErrDesc)
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("PR Document Total " & dAmount, sFuncName)

                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling MatrixDataToDataTable() " & dAmount, sFuncName)
                                        If MatrixDataToDataTable(oForm, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                                        '-----------------/ * Budget Script
                                        p_oDTConsBudget = New DataTable

                                        Sql = "select DocEntry , U_BudName , U_Period ,U_Account , '' [U_OUCode] , U_BUCode , U_PrjCode , U_BudAmount, U_BalAmount, year(FinancYear) + 1 FinancYear  from " & p_sHoldingEntity & " ..[@AB_PROJECTBUDGET] T0 " & _
            "join " & p_sHoldingEntity & " ..[OBGS] T1 on T0.[U_BudName] = T1.[Name] where T1.[U_AB_ACTIVE] = 'Yes'"
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Consolidation Budget " & Sql, sFuncName)
                                        orset.DoQuery(Sql)
                                        p_oDTConsBudget = ConvertRecordset(orset)

                                        '-----------------Budget Script * /

                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Budget_Validation() " & dAmount, sFuncName)
                                        oForm.Freeze(True)
                                        oForm_UDF.Freeze(True)

                                        If Budget_Validation(oForm_UDF, oForm, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                                        sSQL = "SELECT T0.[U_ApprGridCode] FROM " & p_sHoldingEntity & " ..[@AB_APPROVALMATRIX]  T0 WHERE T0.[U_ApprDept]  = '" & sPurchasingDepartment & "' and T0.[U_DocType]  = 'PR' and isnull(T0.[U_PRFROM],0)  <= " & dAmount & " and  (isnull(T0.[U_PRTO],0) >= " & dAmount & " or isnull(T0.[U_PRTO],0) = 0) " & _
                                            "and (cast(isnull(T0.[U_PRFROM],0) as integer) - cast(isnull(T0.[U_PRTO],0) as integer)) <> 0"

                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting Approval Grid Code " & sSQL, sFuncName)
                                        orset.DoQuery(sSQL)
                                        If String.IsNullOrEmpty(orset.Fields.Item("U_ApprGridCode").Value) Then
                                            p_oSBOApplication.StatusBar.SetText("No valid Approval found. Please select another department. ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                            oForm_UDF.Items.Item("U_AB_APPROVALAMT").Specific.String = 0
                                            oForm_UDF.Items.Item("U_AB_APPROVALCODE").Specific.String = String.Empty
                                            oForm.Freeze(False)
                                            oForm_UDF.Freeze(False)
                                            BubbleEvent = False
                                            Exit Try
                                        End If
                                        oForm_UDF.Items.Item("U_AB_APPROVALAMT").Specific.String = dAmount
                                        oForm_UDF.Items.Item("U_AB_APPROVALCODE").Specific.String = orset.Fields.Item("U_ApprGridCode").Value

                                    End If

                                    p_PRApprovalCode = oForm_UDF.Items.Item("U_AB_APPROVALCODE").Specific.String
                                    oForm.Freeze(False)
                                    oForm_UDF.Freeze(False)

                                Catch ex As Exception
                                    oForm.Freeze(False)
                                    oForm_UDF.Freeze(False)
                                    p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                End Try
                                Exit Sub
                            End If
                        End If


                    Case "50103"
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                            Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount(50103, pVal.FormTypeCount)

                            Dim sUser As String = String.Empty
                            Dim sDraftKey As String = String.Empty
                            Dim sQuery As String = String.Empty

                            If pVal.ItemUID = "1" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                Dim oMatrix As SAPbouiCOM.Matrix = Nothing
                                Dim oForm_Mssg As SAPbouiCOM.Form = Nothing
                                Dim orset As SAPbobsCOM.Recordset = Nothing
                                Dim oRow() As Data.DataRow = Nothing

                                Try
                                    'W = Pending
                                    'Y = Approved
                                    'N = Not Approved
                                    orset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oForm_Mssg = p_oSBOApplication.Forms.GetFormByTypeAndCount(198, pVal.FormTypeCount)
                                    oMatrix = oForm_Mssg.Items.Item("6").Specific
                                    If Left(oMatrix.Columns.Item("V_0").Cells.Item(1).Specific.String, 14) = "Purchase Order" Then
                                        If oForm.Items.Item("28").Specific.value = "N" Then
                                            Dim sBody As String = String.Empty
                                            Dim p_SyncDateTime As String = String.Empty
                                            Dim sEmailSubject As String = String.Empty
                                            Dim sUSerName As String = String.Empty

                                            sUser = "%/" & p_oDICompany.UserName & "/%"
                                            sDraftKey = oForm.Items.Item("42").Specific.String

                                            sEmailSubject = "PO Draft No. " & sDraftKey & "  " & p_oDICompany.CompanyName & " has been Rejected "
                                            p_SyncDateTime = Format(Now, "dddd") & ", " & Format(Now, "MMM") & " " & Format(Now, "dd") & ", " & Format(Now, "yyyy") & " " & Format(Now, "HH:mm:ss")
                                            sBody = sBody & "<div align=left style='font-size:10.0pt;font-family:Arial'>"
                                            sBody = sBody & " Dear Sir/Madam,<br /><br />"
                                            sBody = sBody & p_SyncDateTime & " <br /><br />"
                                            sBody = sBody & " " & " <B> Rejected your PO approval in SAP . </B><br /><br />"
                                            sBody = sBody & " " & "<B> PO Draft No. : " & sDraftKey & " </B> (Can be viewed under Main Menu/ Administration/ Approval Procedures/ Approval Status Report) <br />"

                                            oRow = p_oDTUserInformation.Select("USER_CODE='" & p_oDICompany.UserName & "'")
                                            sUSerName = oRow(0).Item("U_NAME").ToString

                                            sBody = sBody & " " & " Doc Rejected by : " & sUSerName & " <br />"
                                            sBody = sBody & " " & " Entity          : " & p_oDICompany.CompanyName
                                            sBody = sBody & " " & " Remarks        : " & oForm.Items.Item("23").Specific.String
                                            sBody = sBody & "<br /><br />"
                                            sBody = sBody & "Thank you."
                                            sBody = sBody & "<br /><br />"
                                            sBody = sBody & " Please do not reply to this email. <div/>"
                                            If String.IsNullOrEmpty(oForm.Items.Item("23").Specific.String) Then
                                                oForm.Items.Item("23").Specific.active = True
                                                p_oSBOApplication.StatusBar.SetText("Remarks should not be Empty .........!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                                BubbleEvent = False
                                                Exit Try
                                            End If

                                            sQuery = "update " & p_sHoldingEntity & " ..[AB_EmailStatus] set [Status] = 'Closed' where " & _
                                             " draftkey = '" & sDraftKey & "' and DocType = 'PO' and Entity = '" & p_oDICompany.CompanyName & "'"
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Changing Status to Closed " & sQuery, sFuncName)
                                            orset.DoQuery(sQuery)

                                            sQuery = "update " & p_sHoldingEntity & " ..[AB_EmailStatus] set [Status] = 'Open', [EmailBody] = '" & Replace(sBody, "'", "''") & "', [EmailSub] = '" & sEmailSubject & "' where seq = " & _
                                               "(select top (1) seq  from " & p_sHoldingEntity & " ..[AB_EmailStatus] where draftkey = '" & sDraftKey & "' and DocType = 'PO' and Entity = '" & p_oDICompany.CompanyName & "' order by cast(Seq as integer) Desc)  and draftkey = '" & sDraftKey & "' and DocType = 'PO' and Entity = '" & p_oDICompany.CompanyName & "'"
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Triggering to Originator " & sQuery, sFuncName)
                                            orset.DoQuery(sQuery)


                                        ElseIf oForm.Items.Item("28").Specific.value = "Y" Then
                                            sUser = "%/" & p_oDICompany.UserName & "/%"
                                            sDraftKey = oForm.Items.Item("42").Specific.String
                                            sQuery = "update " & p_sHoldingEntity & " ..[AB_EmailStatus] set [Status] = 'Open' where seq = " & _
                                                "(select top (1) seq + 1 from " & p_sHoldingEntity & " ..[AB_EmailStatus] where [sUser] like '" & sUser & "' and draftkey = '" & sDraftKey & "' and DocType = 'PO' and Entity = '" & p_oDICompany.CompanyName & "')  and draftkey = '" & sDraftKey & "' and DocType = 'PO' and Entity = '" & p_oDICompany.CompanyName & "' and  [Status]='Pending'"
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Update Next level " & sQuery, sFuncName)
                                            orset.DoQuery(sQuery)
                                        End If
                                    ElseIf Left(oMatrix.Columns.Item("V_0").Cells.Item(1).Specific.String, 16) = "Purchase Request" Then
                                        If oForm.Items.Item("28").Specific.value = "Y" Then
                                            sUser = "%/" & p_oDICompany.UserName & "%/"
                                            sDraftKey = oForm.Items.Item("42").Specific.String
                                            sQuery = "update " & p_sHoldingEntity & " ..[AB_EmailStatus] set [Status] = 'Open' where seq = " & _
                                                "(select top (1) seq + 1 from " & p_sHoldingEntity & " ..[AB_EmailStatus] where [sUser] like '" & sUser & "' and draftkey = '" & sDraftKey & "' and DocType = 'PR' and Entity = '" & p_oDICompany.CompanyName & "') and  draftkey = '" & sDraftKey & "' and DocType = 'PR' and Entity = '" & p_oDICompany.CompanyName & "' and  [Status]='Pending'"
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Update Next level " & sQuery, sFuncName)
                                            orset.DoQuery(sQuery)
                                        ElseIf oForm.Items.Item("28").Specific.value = "N" Then
                                            Dim sBody As String = String.Empty
                                            Dim p_SyncDateTime As String = String.Empty
                                            Dim sEmailSubject As String = String.Empty

                                            sUser = "%/" & p_oDICompany.UserName & "/%"
                                            sDraftKey = oForm.Items.Item("42").Specific.String

                                            sEmailSubject = "PO Draft No. " & sDraftKey & "  " & p_oDICompany.CompanyName & " has been Rejected "
                                            p_SyncDateTime = Format(Now, "dddd") & ", " & Format(Now, "MMM") & " " & Format(Now, "dd") & ", " & Format(Now, "yyyy") & " " & Format(Now, "HH:mm:ss")
                                            sBody = sBody & "<div align=left style='font-size:10.0pt;font-family:Arial'>"
                                            sBody = sBody & " Dear Sir/Madam,<br /><br />"
                                            sBody = sBody & p_SyncDateTime & " <br /><br />"
                                            sBody = sBody & " " & " <B> Rejected your PR approval in SAP . </B><br /><br />"
                                            sBody = sBody & " " & "<B> PR Draft No. : " & sDraftKey & " </B> (Can be viewed under Main Menu/ Administration/ Approval Procedures/ Approval Status Report) <br />"
                                            sBody = sBody & " " & " Doc Rejected by : " & p_oDICompany.UserName & " <br />"
                                            sBody = sBody & " " & " Entity          : " & p_oDICompany.CompanyName
                                            sBody = sBody & " " & " Remarks         : " & oForm.Items.Item("23").Specific.String
                                            sBody = sBody & "<br /><br />"
                                            sBody = sBody & "Thank you."
                                            sBody = sBody & "<br /><br />"
                                            sBody = sBody & " Please do not reply to this email. <div/>"

                                            sQuery = "update " & p_sHoldingEntity & " ..[AB_EmailStatus] set [Status] = 'Closed' where " & _
                                             " draftkey = '" & sDraftKey & "' and DocType = 'PR' and Entity = '" & p_oDICompany.CompanyName & "'"
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Changing Status to Closed " & sQuery, sFuncName)
                                            orset.DoQuery(sQuery)

                                            sQuery = "update " & p_sHoldingEntity & " ..[AB_EmailStatus] set [Status] = 'Open', [EmailBody] = '" & Replace(sBody, "'", "''") & "', [EmailSub] = '" & sEmailSubject & "' where seq = " & _
                                               "(select top (1) seq  from " & p_sHoldingEntity & " ..[AB_EmailStatus] where draftkey = '" & sDraftKey & "' and DocType = 'PR' and Entity = '" & p_oDICompany.CompanyName & "'  order by cast(Seq as integer) Desc)  and draftkey = '" & sDraftKey & "' and DocType = 'PR' and Entity = '" & p_oDICompany.CompanyName & "'"
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Triggering to Originator " & sQuery, sFuncName)
                                            orset.DoQuery(sQuery)


                                        End If
                                    End If

                                Catch ex As Exception
                                    p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                End Try
                                Exit Sub
                            End If
                        End If

                    Case "50105"

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                            Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount(50105, pVal.FormTypeCount)

                            Dim sUser As String = String.Empty
                            Dim sDraftKey As String = String.Empty
                            Dim sQuery As String = String.Empty

                            If pVal.ItemUID = "1" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                Dim oMatrix As SAPbouiCOM.Matrix = Nothing
                                Dim oForm_Mssg As SAPbouiCOM.Form = Nothing
                                Dim orset As SAPbobsCOM.Recordset = Nothing

                                Try
                                    'W = Pending
                                    'Y = Approved
                                    'N = Not Approved
                                    orset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oMatrix = oForm.Items.Item("3").Specific
                                    p_sDocType = String.Empty
                                    For imjs As Integer = 1 To oMatrix.RowCount
                                        If oMatrix.IsRowSelected(imjs) Then
                                            sDraftKey = oMatrix.Columns.Item("540000075").Cells.Item(imjs).Specific.String
                                            p_sDocType = oMatrix.Columns.Item("23").Cells.Item(imjs).Specific.String
                                            Exit For
                                        End If
                                    Next
                                    sUser = "%/" & p_oDICompany.UserName & "/%"
                                    If p_sAstatus = "Y" Then
                                        If p_sDocType = "22" Then
                                            sQuery = "update " & p_sHoldingEntity & " ..[AB_EmailStatus] set [Status] = 'Open' where seq = " & _
                                               "(select top(1) seq + 1 from " & p_sHoldingEntity & " ..[AB_EmailStatus] where [sUser] like '" & sUser & "' and draftkey = '" & sDraftKey & "' and DocType = 'PO' and Entity = '" & p_oDICompany.CompanyName & "' ) and draftkey = '" & sDraftKey & "' and DocType = 'PO' and Entity = '" & p_oDICompany.CompanyName & "' and  [Status]='Pending'"
                                        ElseIf p_sDocType = "1470000113" Then
                                            sQuery = "update " & p_sHoldingEntity & " ..[AB_EmailStatus] set [Status] = 'Open' where seq = " & _
                                               "(select top(1) seq + 1 from " & p_sHoldingEntity & " ..[AB_EmailStatus] where [sUser] like '" & sUser & "' and draftkey = '" & sDraftKey & "' and DocType = 'PR' and Entity = '" & p_oDICompany.CompanyName & "') and  draftkey = '" & sDraftKey & "' and DocType = 'PR' and Entity = '" & p_oDICompany.CompanyName & "' and  [Status]='Pending'"
                                        End If

                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Update Next level " & sQuery, sFuncName)
                                        orset.DoQuery(sQuery)
                                    ElseIf p_sAstatus = "N" Then

                                        Dim sBody As String = String.Empty
                                        Dim p_SyncDateTime As String = String.Empty
                                        Dim sEmailSubject As String = String.Empty
                                        Dim sDoctype As String = String.Empty

                                        If p_sDocType = "22" Then
                                            sDoctype = "PO"
                                        ElseIf p_sDocType = "1470000113" Then
                                            sDoctype = "PR"
                                        End If

                                        sUser = "%/" & p_oDICompany.UserName & "/%"


                                        sEmailSubject = "" & sDoctype & "  Draft No. " & sDraftKey & "  " & p_oDICompany.CompanyName & " has been Rejected "
                                        p_SyncDateTime = Format(Now, "dddd") & ", " & Format(Now, "MMM") & " " & Format(Now, "dd") & ", " & Format(Now, "yyyy") & " " & Format(Now, "HH:mm:ss")
                                        sBody = sBody & "<div align=left style='font-size:10.0pt;font-family:Arial'>"
                                        sBody = sBody & " Dear Sir/Madam,<br /><br />"
                                        sBody = sBody & p_SyncDateTime & " <br /><br />"
                                        sBody = sBody & " " & " <B> Rejected your " & sDoctype & " approval in SAP . </B><br /><br />"
                                        sBody = sBody & " " & "<B> " & sDoctype & " Draft No. : " & sDraftKey & " </B> (Can be viewed under Main Menu/ Administration/ Approval Procedures/ Approval Status Report) <br />"
                                        sBody = sBody & " " & " Doc Rejected by : " & p_oDICompany.UserName & " <br />"
                                        sBody = sBody & " " & " Entity          : " & p_oDICompany.CompanyName

                                        sBody = sBody & "<br /><br />"
                                        sBody = sBody & "Thank you."
                                        sBody = sBody & "<br /><br />"
                                        sBody = sBody & " Please do not reply to this email. <div/>"

                                        sQuery = "update " & p_sHoldingEntity & " ..[AB_EmailStatus] set [Status] = 'Closed' where " & _
                                         "  draftkey = '" & sDraftKey & "' and DocType = '" & sDoctype & "' and Entity = '" & p_oDICompany.CompanyName & "'"
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Changing Status to Closed " & sQuery, sFuncName)
                                        orset.DoQuery(sQuery)

                                        sQuery = "update " & p_sHoldingEntity & " ..[AB_EmailStatus] set [Status] = 'Open', [EmailBody] = '" & Replace(sBody, "'", "''") & "', [EmailSub] = '" & sEmailSubject & "' where seq = " & _
                                           "(select top (1) seq  from " & p_sHoldingEntity & " ..[AB_EmailStatus] where draftkey = '" & sDraftKey & "' and DocType = '" & sDoctype & "' and Entity = '" & p_oDICompany.CompanyName & "' order by cast(Seq as integer) Desc)  and draftkey = '" & sDraftKey & "' and DocType = '" & sDoctype & "' and Entity = '" & p_oDICompany.CompanyName & "'"
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Triggering to Originator " & sQuery, sFuncName)
                                        orset.DoQuery(sQuery)


                                    End If

                                Catch ex As Exception
                                    p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                End Try
                                Exit Sub
                            End If
                        End If

                    Case "142"  ' - --------------------------------PO
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                            Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount(142, pVal.FormTypeCount)
                            p_FormTypecount = pVal.FormTypeCount
                            If pVal.ItemUID = "1" And oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then

                                Dim dAmount As Double = 0
                                Dim oDocDatatable As New DataTable
                                Dim oBaseRefDT As New DataTable
                                Dim oResultDT As New DataTable
                                Dim sBaseRef As String = String.Empty
                                Dim sBaseTable As String = String.Empty
                                Dim oRset As SAPbobsCOM.Recordset = Nothing
                                Dim oRow() As Data.DataRow = Nothing
                                Dim dBAmount As Double = 0
                                Dim sBAccount As String = String.Empty
                                Dim sBCategory As String = String.Empty
                                Dim SQL As String = String.Empty
                                Dim dBalAmount As Double = 0.0
                                Dim DocEntry As String = String.Empty
                                Dim oNewForm As SAPbouiCOM.Form = Nothing
                                Dim dComAmount As Double = 0
                                Dim oComboSeries As SAPbouiCOM.ComboBox


                                Try
                                    sFuncName = "PO Add Event"

                                    oRset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oNewForm = p_oSBOApplication.Forms.GetFormByTypeAndCount("-142", pVal.FormTypeCount)

                                    ''  oNewForm.Items.Item("U_AB_PREAPPROVED").Enabled = False

                                    If oForm.Title = "Purchase Order - Draft [Approved]" Then Exit Try

                                    If oForm.Items.Item("3").Specific.value = "I" Then
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling GetPOMaxAmount() " & SQL, sFuncName)

                                        '--  Getting the doctal of PO
                                        dAmount = GetPOMaxAmount(oForm, sErrDesc)
                                        If dAmount = 0 And sErrDesc.Length > 1 Then Throw New ArgumentException(sErrDesc)

                                        ' oNewForm.Items.Item("U_AB_APPROVALAMT").Specific.value = dAmount


                                        Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("38").Specific
                                        Select Case oMatrix.Columns.Item("43").Cells.Item(1).Specific.String
                                            Case "1470000113"
                                                sBaseTable = "PRQ1"
                                            Case "540000006"
                                                sBaseTable = "PQT1"
                                        End Select
                                        ' BubbleEvent = False
                                        'Exit Sub

                                        p_oDTPOMatrixs.Clear()

                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling  MatrixDataToDataTable() " & SQL, sFuncName)
                                        oDocDatatable = MatrixDataToDataTable(oDocDatatable, oMatrix, oNewForm, "True", sErrDesc)
                                        If Not String.IsNullOrEmpty(sErrDesc) Then
                                            BubbleEvent = False
                                            Exit Sub
                                        End If

                                        '-----------------/ * Budget Script
                                        p_oDTConsBudget = New DataTable

                                        SQL = "select DocEntry , U_BudName , U_Period ,U_Account , '' [U_OUCode] , U_BUCode , U_PrjCode , U_BudAmount, U_BalAmount, year(FinancYear) + 1 FinancYear  from " & p_sHoldingEntity & ".. [@AB_PROJECTBUDGET] T0 " & _
            "join " & p_sHoldingEntity & " ..[OBGS] T1 on T0.[U_BudName] = T1.[Name] where T1.[U_AB_ACTIVE] = 'Yes'"
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Consolidation Budget " & SQL, sFuncName)
                                        oRset.DoQuery(SQL)
                                        p_oDTConsBudget = ConvertRecordset(oRset)

                                        '-----------------Budget Script * /

                                        If String.IsNullOrEmpty(sBaseTable) Then GoTo CostCenterValidation
                                        If oDocDatatable.Rows.Count = 0 Then GoTo CostCenterValidation

                                        oBaseRefDT = oDocDatatable.DefaultView.ToTable(True, "BaseRef")
                                        For imjs As Integer = 0 To oBaseRefDT.Rows.Count - 1
                                            sBaseRef = sBaseRef & oBaseRefDT.Rows(imjs).Item(0).ToString & ","
                                        Next imjs

                                        sBaseRef = Left(Trim(sBaseRef), sBaseRef.Length - 1)

                                        If Not String.IsNullOrEmpty(sBaseRef) Then
                                            SQL = "SELECT T0.[DocEntry], isnull(T0.[LineNum],0) as 'LineNum', T0.[ItemCode], T0.[Dscription], T0.[Quantity], " & _
                                                "T0.[LineTotal] FROM " & sBaseTable & " T0 WHERE T0.[DocEntry] in (" & sBaseRef & ")"
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Base Document " & SQL, sFuncName)
                                            oRset.DoQuery(SQL)
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling  ConvertRecordset() " & SQL, sFuncName)
                                            oResultDT = ConvertRecordset(oRset)
                                            If oResultDT.Rows.Count = 0 Then GoTo CostCenterValidation

                                            '----- Variance validation
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling MarketingDocValidation_PO()", sFuncName)
                                            If MarketingDocValidation_PO(oDocDatatable, oResultDT, oNewForm, sErrDesc) <> RTN_SUCCESS Then
                                                BubbleEvent = False
                                                Exit Sub
                                            End If
                                        End If

CostCenterValidation:

                                        '--- Base Document Check
                                        '' Commented as per Gabriel says
                                        oComboSeries = oForm.Items.Item("88").Specific

                                        If sBaseTable = "PRQ1" Then ''Or oComboSeries.Selected.Description.ToUpper().Trim() = "INTERNAL" Or oComboSeries.Selected.Description.ToUpper().Trim() = "ANNUAL" Then
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Base Document Check - Change the Pre Approval Flag", sFuncName)
                                            p_PREAPPROVED = True
                                            oForm.ActiveItem = "16"
                                            oNewForm.Items.Item("U_AB_PREAPPROVED").Enabled = True
                                            oNewForm.Items.Item("U_AB_PREAPPROVED").Specific.select("Y")
                                            oForm.ActiveItem = "16"
                                            If oComboSeries.Selected.Description.ToUpper().Trim() = "INTERNAL" Or oComboSeries.Selected.Description.ToUpper().Trim() = "ANNUAL" Then
                                            Else
                                                oNewForm.Items.Item("U_AB_PREAPPROVED").Enabled = False
                                            End If

                                        End If


                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Budget_Validation()", sFuncName)
                                        oForm.Freeze(True)
                                        oNewForm.Freeze(True)
                                        '---- Budget Validation  
                                        If Budget_Validation(oNewForm, oForm, sErrDesc) <> RTN_SUCCESS Then
                                            oForm.Freeze(False)
                                            oNewForm.Freeze(False)
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                        p_PREAPPROVED = False
                                        '---- Competitive Quotes Validation will trigger if the amount is >= 10K
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling  POValidation_Competitive() " & SQL, sFuncName)
                                        If POValidation_Competitive(oForm, dAmount, dComAmount, sErrDesc) <> RTN_SUCCESS Then
                                            oForm.Freeze(False)
                                            oNewForm.Freeze(False)
                                            BubbleEvent = False
                                            Exit Sub
                                        End If

                                        '' ''---- Fetch the Approval Grid code

                                        oComboSeries = oForm.Items.Item("88").Specific
                                        '' oComboSeries.Selected.Description.ToUpper().Trim() = "INTERNAL"
                                        If oNewForm.Items.Item("U_AB_PREAPPROVED").Specific.value.ToString.ToUpper.Trim() = "Y" Then
                                            If POValidation_MatrixGridCode_Internal(oForm, dComAmount, sErrDesc) <> RTN_SUCCESS Then
                                                oForm.Freeze(False)
                                                oNewForm.Freeze(False)
                                                BubbleEvent = False
                                                Exit Sub
                                            End If
                                        Else
                                            If POValidation_MatrixGridCode(oForm, dComAmount, sErrDesc) <> RTN_SUCCESS Then
                                                oForm.Freeze(False)
                                                oNewForm.Freeze(False)
                                                BubbleEvent = False
                                                Exit Sub
                                            End If
                                        End If
                                        oForm.Freeze(False)
                                        oNewForm.Freeze(False)
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
                                    End If

                                Catch ex As Exception
                                    oForm.Freeze(False)
                                    oNewForm.Freeze(False)
                                    BubbleEvent = False
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                                    p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    WriteToLogFile(Err.Description, sFuncName)
                                    Exit Sub
                                End Try
                                Exit Sub
                            End If

                        ElseIf pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE Then
                            Approval = False
                            p_POApprovalCode = String.Empty
                            p_PRApprovalCode = String.Empty
                            p_PREAPPROVED = False
                        End If

                    Case "143" '------------------------- GRPO
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                            Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount(143, pVal.FormTypeCount)
                            If pVal.ItemUID = "1" And oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                Try
                                    Dim oDocDatatable As New DataTable
                                    Dim oBaseRefDT As New DataTable
                                    Dim oResultDT As New DataTable
                                    Dim sBaseRef As String = String.Empty
                                    Dim sBaseTable As String = String.Empty
                                    Dim oRset As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                    If oForm.Items.Item("3").Specific.value = "I" Then
                                        Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("38").Specific
                                        Select Case oMatrix.Columns.Item("43").Cells.Item(1).Specific.String
                                            Case "540000006"
                                                sBaseTable = "PQT1"
                                            Case "22"
                                                sBaseTable = "POR1"
                                            Case "21"
                                                sBaseTable = "RPD1"
                                            Case "14"
                                                sBaseTable = "PCH1"
                                        End Select
                                        ' BubbleEvent = False
                                        'Exit Sub

                                        If String.IsNullOrEmpty(sBaseTable) Then Exit Try

                                        oDocDatatable = MatrixDataToDataTable(oDocDatatable, oMatrix, oForm, "False", sErrDesc)
                                        If Not String.IsNullOrEmpty(sErrDesc) Then
                                            BubbleEvent = False
                                            Exit Sub
                                        End If

                                        If oDocDatatable.Rows.Count = 0 Then Exit Sub

                                        oBaseRefDT = oDocDatatable.DefaultView.ToTable(True, "BaseRef")
                                        For imjs As Integer = 0 To oBaseRefDT.Rows.Count - 1
                                            sBaseRef = sBaseRef & oBaseRefDT.Rows(imjs).Item(0).ToString & ","
                                        Next imjs

                                        sBaseRef = Left(Trim(sBaseRef), sBaseRef.Length - 1)

                                        If Not String.IsNullOrEmpty(sBaseRef) Then
                                            Dim SQL As String = "SELECT T0.[DocEntry], isnull(T0.[LineNum],0) as 'LineNum', T0.[ItemCode], T0.[Dscription], T0.[Quantity], " & _
                                                "T0.[LineTotal] FROM " & sBaseTable & " T0 WHERE T0.[DocEntry] in (" & sBaseRef & ")"
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Base Document " & SQL, sFuncName)
                                            oRset.DoQuery(SQL)
                                            oResultDT = ConvertRecordset(oRset)
                                            If oResultDT.Rows.Count = 0 Then Exit Sub

                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DisplayStatus()", sFuncName)
                                            If MarketingDocValidation(oDocDatatable, oResultDT, sErrDesc) <> RTN_SUCCESS Then
                                                BubbleEvent = False
                                                Exit Sub
                                            End If
                                        End If
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
                                    End If

                                Catch ex As Exception
                                    BubbleEvent = False
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                                    p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    WriteToLogFile(Err.Description, sFuncName)
                                    Exit Sub
                                End Try
                                Exit Sub
                            End If
                        End If

                    Case "141" '------------------------- AP Invoice
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                            Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount(141, pVal.FormTypeCount)
                            If pVal.ItemUID = "1" And oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And oForm.Title = "A/P Invoice" Then
                                Try
                                    Dim oDocDatatable As New DataTable
                                    Dim oBaseRefDT As New DataTable
                                    Dim oResultDT As New DataTable
                                    Dim sBaseRef As String = String.Empty
                                    Dim sBaseTable As String = String.Empty
                                    Dim oRset As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                    If oForm.Items.Item("3").Specific.value = "I" Then
                                        Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("38").Specific
                                        Select Case oMatrix.Columns.Item("43").Cells.Item(1).Specific.String
                                            Case "540000006"
                                                sBaseTable = "PQT1"  '' PR
                                            Case "22"
                                                sBaseTable = "POR1"  '' PO
                                            Case "20"
                                                sBaseTable = "PDN1"   ''GRPO
                                        End Select
                                        ' BubbleEvent = False
                                        'Exit Sub
                                        If String.IsNullOrEmpty(sBaseTable) Then Exit Try

                                        oDocDatatable = MatrixDataToDataTable(oDocDatatable, oMatrix, oForm, "False", sErrDesc)
                                        If Not String.IsNullOrEmpty(sErrDesc) Then
                                            BubbleEvent = False
                                            Exit Sub
                                        End If

                                        If oDocDatatable.Rows.Count = 0 Then Exit Sub

                                        oBaseRefDT = oDocDatatable.DefaultView.ToTable(True, "BaseRef")
                                        For imjs As Integer = 0 To oBaseRefDT.Rows.Count - 1
                                            sBaseRef = sBaseRef & oBaseRefDT.Rows(imjs).Item(0).ToString & ","
                                        Next imjs

                                        sBaseRef = Left(Trim(sBaseRef), sBaseRef.Length - 1)

                                        If Not String.IsNullOrEmpty(sBaseRef) Then

                                            If CheckTotal(oForm, oMatrix, sErrDesc) <> RTN_SUCCESS Then
                                                BubbleEvent = False
                                                Exit Sub
                                            End If

                                            ' ''Dim SQL As String = "SELECT T0.[DocEntry], isnull(T0.[LineNum],0) as 'LineNum', T0.[ItemCode], T0.[Dscription], T0.[Quantity], " & _
                                            ' ''    "T0.[LineTotal] FROM " & sBaseTable & " T0 WHERE T0.[DocEntry] in (" & sBaseRef & ")"
                                            ' ''If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Base Document " & SQL, sFuncName)
                                            ' ''oRset.DoQuery(SQL)
                                            ' ''oResultDT = ConvertRecordset(oRset)
                                            ' ''If oResultDT.Rows.Count = 0 Then Exit Sub

                                            ' ''If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DisplayStatus()", sFuncName)
                                            ' ''If MarketingDocValidation(oDocDatatable, oResultDT, sErrDesc) <> RTN_SUCCESS Then
                                            ' ''    BubbleEvent = False
                                            ' ''    Exit Sub
                                            ' ''End If
                                        End If
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
                                    End If

                                Catch ex As Exception
                                    BubbleEvent = False
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                                    p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    WriteToLogFile(Err.Description, sFuncName)
                                    Exit Sub
                                End Try
                                Exit Sub
                            End If
                        End If


                    Case "50106"

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                            Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount(50106, pVal.FormTypeCount)
                            If pVal.ItemUID = "1" And oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                Try
                                    Dim oRset As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim sSQL As String = String.Empty
                                    Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("4").Specific
                                    p_oDTEmailAddress = New DataTable()
                                    Dim sUser As String = String.Empty
                                    Dim sAuser As String = String.Empty
                                    Dim sApprovalGridcode As String = String.Empty
                                    Dim sSplit() As String
                                    Dim iCount As Integer = 0
                                    Dim oRow() As Data.DataRow = Nothing

                                    If Not String.IsNullOrEmpty(p_POApprovalCode) Then
                                        sApprovalGridcode = p_POApprovalCode
                                    Else
                                        sApprovalGridcode = p_PRApprovalCode
                                    End If
                                    ''  sApprovalGridcode = "APPPO333"

                                    If String.IsNullOrEmpty(sApprovalGridcode) Then
                                        p_oSBOApplication.MessageBox("Please close the Approval form and Try again  ")
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Update the Email Address in the below Users " & vbCrLf & sUser, "")
                                        WriteToLogFile("Close the Approval form and Try again ", sFuncName)
                                        BubbleEvent = False
                                        Exit Sub
                                    End If

                                    sSQL = " select char(39) + isnull(T0.U_Appr1,'') + char(39) + ',' + char(39) + isnull(T0.U_Appr2,'') + char(39) + ',' + char(39) + isnull(T0.U_Appr3,'') + char(39) [User]  " & _
                                             " from " & p_sHoldingEntity & " ..[@AB_APPROVALMATRIX] T0 where T0.U_ApprGridCode = '" & sApprovalGridcode & "'"

                                    ''sSQL = " select char(39) + isnull(T0.U_AB_APPROVER_1,'') + char(39) + ',' + char(39) + isnull(T0.U_AB_APPROVER_2,'') + char(39) + ',' + char(39) + isnull(T0.U_AB_APPROVER_3,'') + char(39) [User]  " & _
                                    ''        " from [@AE_APPROVALMATRIX] T0 where T0.U_AB_APPROVALGRIDCOD = '" & sApprovalGridcode & "'"


                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting the Approver from the UDT  " & sSQL, sFuncName)
                                    oRset.DoQuery(sSQL)
                                    sAuser = oRset.Fields.Item("User").Value
                                    sSplit = sAuser.Split(",")
                                    sSQL = String.Empty
                                    iCount = 0
                                    For Each element As String In sSplit
                                        If Not String.IsNullOrEmpty(element) Then
                                            If element <> "''" Then
                                                Dim ss As String = element
                                                oRow = p_oDTUserInformation.Select("USER_CODE = " & element & "")
                                                If oRow.Count = 0 Then
                                                    p_oSBOApplication.MessageBox("Not a Valid User " & vbCrLf & element & vbCrLf & "Approval Grid Code : " & sApprovalGridcode)
                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Not a valid user " & element & " - Approval Grid Code : " & sApprovalGridcode, "")
                                                    WriteToLogFile("Not a valid user " & element & " - Approval Grid Code : " & sApprovalGridcode, sFuncName)
                                                    BubbleEvent = False
                                                    Exit Sub
                                                End If
                                                iCount += 1
                                                If iCount > 1 Then
                                                    sSQL += "Union All"
                                                End If
                                                sSQL += " select  " & iCount & " [SortId], T3.[USER_CODE] [User]," & _
                                          "T3.E_Mail AS [Email], 'Authorizer' [Cat], T3.[U_NAME] [Name] from [OUSR] T3  where [USER_CODE] in" & _
                                          "(" & element & ")"
                                            End If
                                        End If
                                    Next

                                    ''sSQL = " select  1 [SortId], T3.[USER_CODE] [User]," & _
                                    ''  "T3.E_Mail AS [Email], 'Authorizer' [Cat], T3.[U_NAME] [Name] from [OUSR] T3  where [USER_CODE] in" & _
                                    ''  "(" & sSplit(0) & ")"
                                    ''sSQL += "Union All"
                                    ''sSQL += " select  2 [SortId], T3.[USER_CODE] [User]," & _
                                    '' "T3.E_Mail AS [Email], 'Authorizer' [Cat], T3.[U_NAME] [Name] from [OUSR] T3  where [USER_CODE] in" & _
                                    '' "(" & sSplit(1) & ")"
                                    ''sSQL += "Union All"
                                    ''sSQL += " select  3 [SortId], T3.[USER_CODE] [User]," & _
                                    '' "T3.E_Mail AS [Email], 'Authorizer' [Cat], T3.[U_NAME] [Name] from [OUSR] T3  where [USER_CODE] in" & _
                                    '' "(" & sSplit(2) & ")"

                                    If sSQL.Length <= 1 Then
                                        p_oSBOApplication.MessageBox("No Approver Defined in the Approval Matrix Table " & vbCrLf & "Approval Grid Code : " & sApprovalGridcode)
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No Approver Defined in the Approval Matrix Table " & "Approval Grid Code : " & sApprovalGridcode, "")
                                        WriteToLogFile("No Approver defined in the Approval matric table " & "Approval Grid Code " & sApprovalGridcode, sFuncName)
                                        BubbleEvent = False
                                        Exit Sub
                                    End If

                                    sSQL += "Union All"
                                    sSQL += " select  4 [SortId], T3.[USER_CODE] [User]," & _
                                     "T3.E_Mail AS [Email], 'Originator' [Cat], T3.[U_NAME] [Name] from [OUSR] T3  where [USER_CODE] in" & _
                                     "('" & p_oDICompany.UserName & "')"

                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting the Email ID from the User setup  " & sSQL, sFuncName)
                                    oRset.DoQuery(sSQL)

                                    p_oDTEmailAddress = ConvertRecordset(oRset)
                                    For Each odr As DataRow In p_oDTEmailAddress.Rows
                                        If String.IsNullOrEmpty(odr("Email").ToString.Trim) Then
                                            sUser += odr("User").ToString.Trim() & " "
                                        End If
                                    Next
                                    If sUser.Length > 0 Then
                                        p_oSBOApplication.MessageBox("Update the Email Address in the below Users " & vbCrLf & sUser & vbCrLf & "To Proceed")
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Update the Email Address in the below Users " & vbCrLf & sUser, "")
                                        WriteToLogFile("Update the Email Address in the below Users " & vbCrLf & sUser, sFuncName)
                                        BubbleEvent = False
                                        Exit Sub
                                    End If


                                    '' p_sEmailAddress = oRset.Fields.Item("Email").Value
                                    ''  If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(" Authorizer Email Address " & p_sEmailAddress, sFuncName)
                                    Approval = True

                                Catch ex As Exception
                                    BubbleEvent = False
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                                    p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                    WriteToLogFile(Err.Description, "")
                                    Exit Sub
                                End Try


                            End If

                            If pVal.ItemUID = "2" And oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                Approval = False
                                p_PRApprovalCode = String.Empty
                                p_POApprovalCode = String.Empty
                            End If

                        End If

                        ' Case "392", "393"
                    Case "3002"

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                            If pVal.ItemUID = "3" Then
                                Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
                                Try
                                    Dim oMAtrix As SAPbouiCOM.Matrix = Nothing
                                    oMAtrix = oForm.Items.Item("3").Specific
                                    p_sStatus = oMAtrix.Columns.Item("50").Cells.Item(pVal.Row).Specific.String
                                    p_sAppStatus = oMAtrix.Columns.Item("51").Cells.Item(pVal.Row).Specific.String
                                Catch ex As Exception
                                    p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    Call WriteToLogFile(sErrDesc, sFuncName)
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
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
            p_oSBOApplication.StatusBar.SetText(sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            WriteToLogFile(Err.Description, sFuncName)
            ShowErr(sErrDesc)
        End Try

    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Sub SBO_Application_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles SBO_Application.FormDataEvent

        Dim oForm As SAPbouiCOM.Form = Nothing

        If BusinessObjectInfo.ActionSuccess = True And BusinessObjectInfo.BeforeAction = False Then

            If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Then
                Select Case BusinessObjectInfo.FormTypeEx

                    '----- Commented BP Approval as per the NIK advise

                    'Case "134"
                    '    Dim oRset As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '    Dim sQuery As String = String.Empty
                    '    Dim dAmount As Double = 0
                    '    Dim sEmailsender As String = String.Empty
                    '    Dim sInsertSQL As String = String.Empty
                    '    Dim sEmailSubject As String = String.Empty
                    '    Dim sBody As String = String.Empty
                    '    Dim p_SyncDateTime As String = String.Empty
                    '    Dim sSQLstring As String = String.Empty
                    '    Dim sDraftkey As String = String.Empty
                    '    Dim sBPCode As String = String.Empty
                    '    Dim sBPName As String = String.Empty


                    '    Try
                    '        oForm = p_oSBOApplication.Forms.GetFormByTypeAndCount(BusinessObjectInfo.FormTypeEx, p_FormTypecount)
                    '        Dim oForm_UDF As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount(-134, p_FormTypecount)
                    '        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("BP FormDataEvent - ADD Mode ", sFuncName)
                    '        If oForm_UDF.Items.Item("U_AB_STATUS").Specific.value.ToString.Trim() = "PENDING" Then
                    '            sBPCode = oForm.DataSources.DBDataSources.Item(0).GetValue("CardCode", 0).ToString.Trim()
                    '            sBPName = oForm.DataSources.DBDataSources.Item(0).GetValue("CardName", 0).ToString.Trim()

                    '            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("status is Pending ", sFuncName)

                    '            sEmailSubject = "Request for BP approval in SAP - " & p_oDICompany.CompanyName & "  BP Code. " & sBPCode
                    '            p_SyncDateTime = Format(Now, "dddd") & ", " & Format(Now, "MMM") & " " & Format(Now, "dd") & ", " & Format(Now, "yyyy") & " " & Format(Now, "HH:mm:ss")
                    '            sBody = sBody & "<div align=left style='font-size:10.0pt;font-family:Arial'>"
                    '            sBody = sBody & " Dear Sir/Madam,<br /><br />"
                    '            sBody = sBody & p_SyncDateTime & " <br /><br />"
                    '            sBody = sBody & " " & " <B> Request for your Business Partner approval in SAP . </B><br /><br />"
                    '            sBody = sBody & " " & "<B> BP Code      : " & sBPCode & " </B> (Can be viewed under Main Menu/ Business Partners) <br />"
                    '            sBody = sBody & " " & " BP name         : " & sBPName & " <br />"
                    '            sBody = sBody & " " & " Originator      : " & p_oDICompany.UserName & " <br />"
                    '            sBody = sBody & " " & " Entity          : " & p_oDICompany.CompanyName

                    '            sBody = sBody & "<br /><br />"
                    '            sBody = sBody & " Please login to the above mentioned entity to approve the document."
                    '            sBody = sBody & "<br /><br /> "
                    '            sBody = sBody & "Thank you."
                    '            sBody = sBody & "<br /><br />"
                    '            sBody = sBody & " Please do not reply to this email. <div/>"

                    '            sInsertSQL += "INSERT INTO " & p_sHoldingEntity & ".. [AB_EmailStatus]  " & _
                    '    "([DocType] , [ObjectType] ,[Entity] , [EmailBody], [EmailSub],  [EmailID] , [Status], [sUser],[Seq],[DraftKey] ) " & _
                    '    " VALUES('BP' , '2', '" & p_oDICompany.CompanyName & "', '" & Replace(sBody, "'", "''") & "', '" & sEmailSubject & "' ,'" & p_oCompDef.sApprover.Trim() & "', 'Open', " & _
                    '    "'" & p_oDICompany.UserName & "'," & 1 & ", '0' ) "

                    '            If sInsertSQL.Length > 0 Then
                    '                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Insert in EmailStatus Table " & sInsertSQL, sFuncName)
                    '                oRset.DoQuery(sInsertSQL)
                    '            End If

                    '        End If


                    '    Catch ex As Exception
                    '        p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    '        Call WriteToLogFile(sErrDesc, sFuncName)
                    '        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                    '    Finally
                    '        oRset = Nothing
                    '    End Try

                    Case "142"
                        Dim oRset As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim sQuery As String = String.Empty
                        Dim dAmount As Double = 0
                        Dim sEmailsender As String = String.Empty
                        Dim sInsertSQL As String = String.Empty
                        Dim sEmailSubject As String = String.Empty
                        Dim sBody As String = String.Empty
                        Dim p_SyncDateTime As String = String.Empty
                        Dim sSQLstring As String = String.Empty
                        Dim sDraftkey As String = String.Empty

                        Try
                            sFuncName = "PO Form Data Event"
                            oForm = p_oSBOApplication.Forms.GetFormByTypeAndCount(142, p_FormTypecount)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Purchase Order ", sFuncName)

                            ''For Each dr As DataRow In p_oDTPOMatrixs.Rows
                            ''    sInsertSQL = "update " & p_sHoldingEntity & ".. [@AB_PROJECTBUDGET] set [U_BalAmount] = '" & CDbl(dr.Item("UpdateAmount").ToString.Trim()) & "' where DocEntry = '" & dr.Item("DocEntry").ToString.Trim() & "'"
                            ''    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Budget Table Insert SQL " & sInsertSQL, sFuncName)
                            ''    oRset.DoQuery(sInsertSQL)
                            ''Next

                            sInsertSQL = String.Empty

                            If Approval = True Then
                                Dim sBPcode As String = String.Empty
                                Dim dDoctotal As Double = 0
                                Dim sOrginator As String = String.Empty
                                Dim sUserName As String = String.Empty
                                Dim sUserNameA As String = String.Empty
                                Dim sComments As String = String.Empty
                                Dim dDoctotalLC As String = 0.0
                                Dim dDoctatalFC As String = 0.0
                                Dim sCurrency As String = String.Empty

                                Dim icount As Integer = 0
                                Dim oDatarow() As DataRow

                                '' Dim oUser As SAPbobsCOM.Users = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUsers)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Approval flag is True ", sFuncName)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Approvel level " & p_oDTEmailAddress.Rows.Count, sFuncName)
                                If p_oDTEmailAddress.Rows.Count > 0 Then
                                    sBPcode = oForm.DataSources.DBDataSources.Item(0).GetValue("CardName", 0).ToString.Trim()
                                    dDoctotal = oForm.DataSources.DBDataSources.Item(0).GetValue("DocTotal", 0)
                                    sComments = oForm.DataSources.DBDataSources.Item(0).GetValue("Comments", 0).ToString.Trim()
                                    oDatarow = p_oDTEmailAddress.Select("Cat='Originator'")
                                    sOrginator = oDatarow(0).Item("Name").ToString
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("1", sFuncName)
                                    sSQLstring = "select max(cast(DocEntry as integer)) [Draftkey] from ODRF where ODRF.ObjType = '22'"
                                    oRset.DoQuery(sSQLstring)
                                    sDraftkey = oRset.Fields.Item("Draftkey").Value

                                    sSQLstring = "SELECT Currency , cast( cast(sum(T0.[LineTotal]) as decimal(18,2)) as nvarchar) [LC], cast(cast(sum(T0.[TotalFrgn]) as decimal(18,2)) as nvarchar) [FC] FROM DRF1 T0 WHERE T0.[DocEntry] = '" & sDraftkey & "' group by  Currency"
                                    oRset.DoQuery(sSQLstring)
                                    dDoctatalFC = oRset.Fields.Item("FC").Value
                                    dDoctotalLC = oRset.Fields.Item("LC").Value
                                    sCurrency = oRset.Fields.Item("Currency").Value

                                    For Each odr As DataRow In p_oDTEmailAddress.Rows
                                        sUserName += odr("Name").ToString.Trim() & ","
                                        icount += 1
                                        sBody = String.Empty

                                        If icount = 1 And odr("Cat").ToString.Trim() = "Authorizer" Then
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Authorizer " & sUserName, sFuncName)
                                            sEmailSubject = "Request for PO approval in SAP - " & p_oDICompany.CompanyName & " / PO Draft No. " & sDraftkey
                                            p_SyncDateTime = Format(Now, "dddd") & ", " & Format(Now, "MMM") & " " & Format(Now, "dd") & ", " & Format(Now, "yyyy") & " " & Format(Now, "HH:mm:ss")
                                            sBody = sBody & "<div align=left style='font-size:10.0pt;font-family:Arial'>"
                                            sBody = sBody & " Dear Sir/Madam,<br /><br />"
                                            sBody = sBody & p_SyncDateTime & " <br /><br />"
                                            sBody = sBody & " " & " <B> Request for PO approval in SAP . </B><br /><br />"
                                            '' sBody = sBody & " " & "<B> PO Draft No. : " & sDraftkey & " </B> (Can be viewed under Main Menu/ Administration/ Approval Procedures/ Approval Status Report) <br />"
                                            sBody = sBody & " " & "<B> PO Draft No. : " & sDraftkey & " </B> (Can be viewed under > Main Menu > Modules > Approval > Approval Window in any entity with the new enhancement) <br />"
                                            sBody = sBody & "<br /><br />"
                                            sBody = sBody & " " & " BP name         : " & sBPcode & " <br />"
                                            sBody = sBody & " " & " Document Amount before GST : SGD " & dDoctotalLC & " <br />"
                                            If CDbl(dDoctatalFC) > 0 Then
                                                sBody = sBody & " " & " Document Amount before GST : " & sCurrency & " " & dDoctatalFC & " <br />"
                                            End If
                                            sBody = sBody & " " & " Originator      : " & sOrginator & " <br />"
                                            sBody = sBody & " " & " Doc approved by : Nil (You are the first approver)" & " <br />"
                                            sBody = sBody & " " & " Entity          : " & p_oDICompany.CompanyName & " <br />"
                                            sBody = sBody & " " & " Remarks         : " & sComments
                                            sBody = sBody & "<br /><br />"
                                            '' sBody = sBody & " Please login to the above mentioned entity to approve the document."
                                            sBody = sBody & " Please login to SAP to approve the document."
                                            sBody = sBody & "<br /><br />"
                                            sBody = sBody & " You may also approve via the old method, that is Main Menu > Modules >  Administration > Approval Procedures> Approval Status Report but you will need to log into the specific entity which the document is created."
                                            sBody = sBody & "<br /><br />"
                                            sBody = sBody & "Thank you."
                                            sBody = sBody & "<br /><br />"
                                            sBody = sBody & " Please do not reply to this email. <div/>"

                                            sInsertSQL += "INSERT INTO " & p_sHoldingEntity & ".. [AB_EmailStatus]  " & _
                                    "([DocType] , [ObjectType] ,[Entity] , [EmailBody], [EmailSub],  [EmailID] , [Status], [sUser],[Seq],[DraftKey] ) " & _
                                    " VALUES('PO' , '22', '" & p_oDICompany.CompanyName & "', '" & Replace(sBody, "'", "''") & "', '" & sEmailSubject & "' ,'" & odr("Email").ToString.Trim() & "', 'Open', " & _
                                    "'/" & Left(odr("User").ToString.Trim(), 28) & "/'," & icount & ", '" & sDraftkey & "' ) "
                                        ElseIf odr("Cat").ToString.Trim() = "Authorizer" Then
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Authorizer 2 ", sFuncName)
                                            Dim susersplit() As String
                                            susersplit = sUserName.Split(",")
                                            sUserNameA = String.Empty
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Split Lenght " & susersplit.Length, sFuncName)
                                            For imjs As Integer = 0 To susersplit.Length - 3
                                                sUserNameA += susersplit(imjs).ToString & ","
                                            Next
                                            sUserNameA = Left(sUserNameA, sUserNameA.Length - 1)
                                            sEmailSubject = "Request for PO approval in SAP - " & p_oDICompany.CompanyName & " / PO Draft No. " & sDraftkey
                                            p_SyncDateTime = Format(Now, "dddd") & ", " & Format(Now, "MMM") & " " & Format(Now, "dd") & ", " & Format(Now, "yyyy") & " " & Format(Now, "HH:mm:ss")
                                            sBody = sBody & "<div align=left style='font-size:10.0pt;font-family:Arial'>"
                                            sBody = sBody & " Dear Sir/Madam,<br /><br />"
                                            sBody = sBody & p_SyncDateTime & " <br /><br />"
                                            sBody = sBody & " " & "<B> Request for PO approval in SAP . </B><br /><br />"
                                            '' sBody = sBody & " " & "<B> PO Draft No. : " & sDraftkey & " </B> (Can be viewed under Main Menu/ Administration/ Approval Procedures/ Approval Status Report) <br />"
                                            sBody = sBody & " " & "<B> PO Draft No. : " & sDraftkey & " </B> (Can be viewed under > Main Menu > Modules > Approval > Approval Window in any entity with the new enhancement) <br />"
                                            sBody = sBody & "<br /><br />"
                                            sBody = sBody & " " & " BP name         : " & sBPcode & " <br />"
                                            sBody = sBody & " " & " Document Amount before GST : SGD " & dDoctotalLC & " <br />"
                                            If CDbl(dDoctatalFC) > 0 Then
                                                sBody = sBody & " " & " Document Amount before GST : " & sCurrency & " " & dDoctatalFC & " <br />"
                                            End If
                                            sBody = sBody & " " & " Originator      : " & sOrginator & " <br />"
                                            sBody = sBody & " " & " Doc approved by : " & sUserNameA & " <br />"
                                            sBody = sBody & " " & " Entity          : " & p_oDICompany.CompanyName & "<br />"
                                            sBody = sBody & " " & " Remarks        : " & sComments
                                            sBody = sBody & "<br /><br />"
                                            '' sBody = sBody & " Please login to the above mentioned entity to approve the document."
                                            sBody = sBody & " Please login to SAP to approve the document."
                                            sBody = sBody & "<br /><br />"
                                            sBody = sBody & " You may also approve via the old method, that is Main Menu > Modules >  Administration > Approval Procedures> Approval Status Report but you will need to log into the specific entity which the document is created."

                                            sBody = sBody & "<br /><br />"
                                            sBody = sBody & "Thank you."
                                            sBody = sBody & "<br /><br />"
                                            sBody = sBody & " Please do not reply to this email. <div/>"

                                            sInsertSQL += "INSERT INTO " & p_sHoldingEntity & ".. [AB_EmailStatus]  " & _
                                     "([DocType] , [ObjectType] ,[Entity] , [EmailBody], [EmailSub],  [EmailID] , [Status], [sUser],[Seq],[DraftKey] ) " & _
                                     " VALUES('PO' , '22', '" & p_oDICompany.CompanyName & "', '" & Replace(sBody, "'", "''") & "', '" & sEmailSubject & "' ,'" & odr("Email").ToString.Trim() & "', 'Pending', " & _
                                     "'/" & Left(odr("User").ToString.Trim(), 28) & "/'," & icount & ", '" & sDraftkey & "' ) "
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("9  " & sInsertSQL, sFuncName)
                                        ElseIf odr("Cat").ToString.Trim() = "Originator" Then
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Originator  ", sFuncName)
                                            sEmailSubject = "PO Draft No. " & sDraftkey & "  " & p_oDICompany.CompanyName & " has been approved for your action "
                                            p_SyncDateTime = Format(Now, "dddd") & ", " & Format(Now, "MMM") & " " & Format(Now, "dd") & ", " & Format(Now, "yyyy") & " " & Format(Now, "HH:mm:ss")
                                            sBody = sBody & "<div align=left style='font-size:10.0pt;font-family:Arial'>"
                                            sBody = sBody & " Dear Sir/Madam,<br /><br />"
                                            sBody = sBody & p_SyncDateTime & " <br /><br />"
                                            sBody = sBody & " " & "<B> The above PO has been approved in SAP for your final action </B><br /><br />"
                                            sBody = sBody & " " & "<B> PO Draft (Approved) No. : " & sDraftkey & " </B> (Can be viewed under Main Menu/ Administration/ Approval Procedures/ Approval Status Report) <br />"
                                            sBody = sBody & " " & " BP name         : " & sBPcode & " <br />"
                                            sBody = sBody & " " & " Document Amount before GST : SGD " & dDoctotalLC & " <br />"
                                            If CDbl(dDoctatalFC) > 0 Then
                                                sBody = sBody & " " & " Document Amount before GST : " & sCurrency & " " & dDoctatalFC & " <br />"
                                            End If
                                            sBody = sBody & " " & " Entity          : " & p_oDICompany.CompanyName & "<br />"
                                            sBody = sBody & " " & " Remarks        : " & sComments
                                            sBody = sBody & "<br /><br />"
                                            sBody = sBody & " Please login to " & p_oDICompany.CompanyName & " and convert the PO Draft(Approved) to PO(Approved) document by clicking ""ADD"" "
                                            sBody = sBody & "<br /><br />"
                                            sBody = sBody & "Thank you."
                                            sBody = sBody & "<br /><br />"
                                            sBody = sBody & " Please do not reply to this email. <div/>"

                                            sInsertSQL += "INSERT INTO " & p_sHoldingEntity & ".. [AB_EmailStatus]  " & _
                                    "([DocType] , [ObjectType] ,[Entity] , [EmailBody], [EmailSub],  [EmailID] , [Status], [sUser],[Seq],[DraftKey] ) " & _
                                    " VALUES('PO' , '22', '" & p_oDICompany.CompanyName & "', '" & Replace(sBody, "'", "''") & "', '" & sEmailSubject & "' ,'" & odr("Email").ToString.Trim() & "', 'Pending', " & _
                                    "'/" & Left(odr("User").ToString.Trim(), 28) & "/'," & icount & ", '" & sDraftkey & "' ) "

                                        End If
                                    Next

                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Email Table Insert SQL " & sInsertSQL, sFuncName)
                                    If sInsertSQL.Length > 0 Then
                                        oRset.DoQuery(sInsertSQL)
                                    End If

                                Else
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Authorizer Email address is Empty ", sFuncName)
                                End If
                                Approval = False
                            End If

                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

                        Catch ex As Exception
                            p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                        Finally
                            oRset = Nothing

                        End Try

                    Case "1470000200"

                        Dim oRset As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim sQuery As String = String.Empty
                        Dim dAmount As Double = 0
                        Dim sEmailsender As String = String.Empty
                        Dim sInsertSQL As String = String.Empty
                        Dim sEmailSubject As String = String.Empty
                        Dim sBody As String = String.Empty
                        Dim p_SyncDateTime As String = String.Empty
                        Dim sSQLstring As String = String.Empty
                        Dim sDraftkey As String = String.Empty
                        Dim sCreator As String = String.Empty
                        Dim sComments As String = String.Empty
                        Dim dDoctotalLC As String = 0.0
                        Dim dDoctatalFC As String = 0.0
                        Dim sCurrency As String = String.Empty
                        Dim sPRNo As String = String.Empty

                        Try
                            sFuncName = "PR Form Data Event"
                            oForm = p_oSBOApplication.Forms.GetFormByTypeAndCount(1470000200, p_FormTypecount)

                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Purchase Request ", sFuncName)

                            If Approval = True Then
                                Dim sBPcode As String = String.Empty
                                Dim dDoctotal As Double = 0
                                Dim sOrginator As String = String.Empty
                                Dim sUserName As String = String.Empty
                                Dim icount As Integer = 0
                                Dim oDatarow() As DataRow
                                Dim sUserNameA As String = String.Empty


                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Approval flag is True ", sFuncName)
                                If p_oDTEmailAddress.Rows.Count > 0 Then

                                    sBPcode = oForm.DataSources.DBDataSources.Item(0).GetValue("CardName", 0).ToString.Trim()
                                    dDoctotal = oForm.DataSources.DBDataSources.Item(0).GetValue("DocTotal", 0)
                                    sComments = oForm.DataSources.DBDataSources.Item(0).GetValue("Comments", 0)
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Originator", sFuncName)
                                    oDatarow = p_oDTEmailAddress.Select("Cat='Originator'")
                                    sOrginator = oDatarow(0).Item("Name").ToString
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Originator " & sOrginator, sFuncName)
                                    sSQLstring = "select max(cast(DocEntry as integer)) [Draftkey] from ODRF where ODRF.ObjType = '1470000113'"
                                    oRset.DoQuery(sSQLstring)
                                    sDraftkey = oRset.Fields.Item("Draftkey").Value

                                    sSQLstring = "SELECT Currency , cast( cast(sum(T0.[LineTotal]) as decimal(18,2)) as nvarchar) [LC], cast(cast(sum(T0.[TotalFrgn]) as decimal(18,2)) as nvarchar) [FC] FROM DRF1 T0 WHERE T0.[DocEntry] = '" & sDraftkey & "' group by  Currency"
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting Document Total  " & sSQLstring, sFuncName)
                                    oRset.DoQuery(sSQLstring)
                                    dDoctatalFC = oRset.Fields.Item("FC").Value
                                    dDoctotalLC = oRset.Fields.Item("LC").Value
                                    sCurrency = oRset.Fields.Item("Currency").Value

                                    For Each odr As DataRow In p_oDTEmailAddress.Rows
                                        sUserName += odr("Name").ToString.Trim() & ","
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("sUserName " & sUserName, sFuncName)
                                        icount += 1
                                        sBody = String.Empty
                                        If icount = 1 And odr("Cat").ToString.Trim() = "Authorizer" Then
                                            '
                                            sEmailSubject = "Request for PR approval in SAP - " & p_oDICompany.CompanyName & " / PR Draft No. " & sDraftkey
                                            p_SyncDateTime = Format(Now, "dddd") & ", " & Format(Now, "MMM") & " " & Format(Now, "dd") & ", " & Format(Now, "yyyy") & " " & Format(Now, "HH:mm:ss")
                                            sBody = sBody & "<div align=left style='font-size:10.0pt;font-family:Arial'>"
                                            sBody = sBody & " Dear Sir/Madam,<br /><br />"
                                            sBody = sBody & p_SyncDateTime & " <br /><br />"
                                            sBody = sBody & " " & " <B> Request for PR approval in SAP . </B><br /><br />"
                                            '' sBody = sBody & " " & "<B> PR Draft No. : " & sDraftkey & " </B> (Can be viewed under Main Menu/ Administration/ Approval Procedures/ Approval Status Report) <br />"
                                            sBody = sBody & " " & "<B> PR Draft No. : " & sDraftkey & " </B> (Can be viewed under > Main Menu > Modules > Approval > Approval Window in any entity with the new enhancement) <br />"
                                            sBody = sBody & "<br /><br />"
                                            sBody = sBody & " " & " Document Amount before GST : SGD " & dDoctotalLC & " <br />"
                                            If CDbl(dDoctatalFC) > 0 Then
                                                sBody = sBody & " " & " Document Amount before GST : " & sCurrency & " " & dDoctatalFC & " <br />"
                                            End If
                                            sBody = sBody & " " & " Originator      : " & sOrginator & " <br />"
                                            sBody = sBody & " " & " Doc approved by : Nil (You are the first approver)" & " <br />"
                                            sBody = sBody & " " & " Entity          : " & p_oDICompany.CompanyName & " <br />"
                                            sBody = sBody & " " & " Remarks        : " & sComments
                                            sBody = sBody & "<br /><br />"
                                            '' sBody = sBody & " Please login to the above mentioned entity to approve the document."
                                            sBody = sBody & " Please login to SAP to approve the document."
                                            sBody = sBody & "<br /><br />"
                                            sBody = sBody & " You may also approve via the old method, that is Main Menu > Modules >  Administration > Approval Procedures> Approval Status Report but you will need to log into the specific entity which the document is created."

                                            sBody = sBody & "<br /><br />"
                                            sBody = sBody & "Thank you."
                                            sBody = sBody & "<br /><br />"
                                            sBody = sBody & " Please do not reply to this email. <div/>"

                                            sInsertSQL += "INSERT INTO " & p_sHoldingEntity & ".. [AB_EmailStatus]  " & _
                                    "([DocType] , [ObjectType] ,[Entity] , [EmailBody], [EmailSub],  [EmailID] , [Status], [sUser],[Seq],[DraftKey] ) " & _
                                    " VALUES('PR' , '1470000113', '" & p_oDICompany.CompanyName & "', '" & Replace(sBody, "'", "''") & "', '" & sEmailSubject & "' ,'" & odr("Email").ToString.Trim() & "', 'Open', " & _
                                    "'/" & Left(odr("User").ToString.Trim(), 28) & "/'," & icount & ", '" & sDraftkey & "' ) "
                                        ElseIf odr("Cat").ToString.Trim() = "Authorizer" Then

                                            Dim susersplit() As String
                                            susersplit = sUserName.Split(",")
                                            sUserNameA = String.Empty
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Split Lenght " & susersplit.Length, sFuncName)
                                            For imjs As Integer = 0 To susersplit.Length - 3
                                                sUserNameA += susersplit(imjs).ToString & ","
                                            Next
                                            sUserNameA = Left(sUserNameA, sUserNameA.Length - 1)

                                            ''Dim susersplit() As String
                                            ''sUserName = Left(sUserName, sUserName.Length - 1)
                                            ''susersplit = sUserName.Split(",")
                                            ''sUserName = String.Empty
                                            ''For imjs As Integer = 0 To susersplit.Length - 2
                                            ''    sUserName += susersplit(imjs).ToString & ","
                                            ''Next
                                            ''sUserName = Left(sUserName, sUserName.Length - 1)
                                            sEmailSubject = "Request for PR approval in SAP - " & p_oDICompany.CompanyName & " / PR Draft No. " & sDraftkey
                                            p_SyncDateTime = Format(Now, "dddd") & ", " & Format(Now, "MMM") & " " & Format(Now, "dd") & ", " & Format(Now, "yyyy") & " " & Format(Now, "HH:mm:ss")
                                            sBody = sBody & "<div align=left style='font-size:10.0pt;font-family:Arial'>"
                                            sBody = sBody & " Dear Sir/Madam,<br /><br />"
                                            sBody = sBody & p_SyncDateTime & " <br /><br />"
                                            sBody = sBody & " " & "<B> Request for PR approval in SAP . </B><br /><br />"
                                            'sBody = sBody & " " & "<B> PR Draft No. : " & sDraftkey & " </B> (Can be viewed under Main Menu/ Administration/ Approval Procedures/ Approval Status Report) <br />"
                                            sBody = sBody & " " & "<B> PR Draft No. : " & sDraftkey & " </B> (Can be viewed under > Main Menu > Modules > Approval > Approval Window in any entity with the new enhancement) <br />"
                                            sBody = sBody & "<br /><br />"
                                            sBody = sBody & " " & " Document Amount before GST : SGD " & dDoctotalLC & " <br />"
                                            If CDbl(dDoctatalFC) > 0 Then
                                                sBody = sBody & " " & " Document Amount before GST : " & sCurrency & " " & dDoctatalFC & " <br />"
                                            End If
                                            sBody = sBody & " " & " Originator      : " & sOrginator & " <br />"
                                            sBody = sBody & " " & " Doc approved by : " & sUserNameA & " <br />"
                                            sBody = sBody & " " & " Entity          : " & p_oDICompany.CompanyName & " <br />"
                                            sBody = sBody & " " & " Remarks        : " & sComments
                                            sBody = sBody & "<br /><br />"
                                            'sBody = sBody & " Please login to the above mentioned entity to approve the document."
                                            sBody = sBody & " Please login to SAP to approve the document."
                                            sBody = sBody & "<br /><br />"
                                            sBody = sBody & " You may also approve via the old method, that is Main Menu > Modules >  Administration > Approval Procedures> Approval Status Report but you will need to log into the specific entity which the document is created."

                                            sBody = sBody & "<br /><br />"
                                            sBody = sBody & "Thank you."
                                            sBody = sBody & "<br /><br />"
                                            sBody = sBody & " Please do not reply to this email. <div/>"

                                            sInsertSQL += "INSERT INTO " & p_sHoldingEntity & ".. [AB_EmailStatus]  " & _
                                     "([DocType] , [ObjectType] ,[Entity] , [EmailBody], [EmailSub],  [EmailID] , [Status], [sUser],[Seq],[DraftKey] ) " & _
                                     " VALUES('PR' , '1470000113', '" & p_oDICompany.CompanyName & "', '" & Replace(sBody, "'", "''") & "', '" & sEmailSubject & "' ,'" & odr("Email").ToString.Trim() & "', 'Pending', " & _
                                     "'/" & Left(odr("User").ToString.Trim(), 28) & "/'," & icount & ", '" & sDraftkey & "' ) "
                                        ElseIf odr("Cat").ToString.Trim() = "Originator" Then
                                            sEmailSubject = "PR Draft No. " & sDraftkey & "  " & p_oDICompany.CompanyName & " has been approved for your action "
                                            p_SyncDateTime = Format(Now, "dddd") & ", " & Format(Now, "MMM") & " " & Format(Now, "dd") & ", " & Format(Now, "yyyy") & " " & Format(Now, "HH:mm:ss")
                                            sBody = sBody & "<div align=left style='font-size:10.0pt;font-family:Arial'>"
                                            sBody = sBody & " Dear Sir/Madam,<br /><br />"
                                            sBody = sBody & p_SyncDateTime & " <br /><br />"
                                            sBody = sBody & " " & "<B> The above PR has been approved in SAP for your final action </B><br /><br />"
                                            sBody = sBody & " " & "<B> PR Draft (Approved) No. : " & sDraftkey & " </B> (Can be viewed under Main Menu/ Administration/ Approval Procedures/ Approval Status Report) <br />"
                                            sBody = sBody & "<br /><br />"
                                            sBody = sBody & " " & " Document Amount before GST : SGD " & dDoctotalLC & " <br />"
                                            If CDbl(dDoctatalFC) > 0 Then
                                                sBody = sBody & " " & " Document Amount before GST : " & sCurrency & " " & dDoctatalFC & " <br />"
                                            End If
                                            sBody = sBody & " " & " Entity          : " & p_oDICompany.CompanyName & " <br />"
                                            sBody = sBody & " " & " Remarks        : " & sComments
                                            sBody = sBody & "<br /><br />"
                                            sBody = sBody & " Please login to the above mentioned entity to convert the PR Draft(Approved) to PR(Approved) document by clicking ""ADD"". The PR will "
                                            sBody = sBody & " then be auto sent to purchasing Unit as stated on the PR."
                                            sBody = sBody & "<br /><br />"
                                            sBody = sBody & "Thank you."
                                            sBody = sBody & "<br /><br />"
                                            sBody = sBody & " Please do not reply to this email. <div/>"


                                            sInsertSQL += "INSERT INTO " & p_sHoldingEntity & ".. [AB_EmailStatus]  " & _
                                    "([DocType] , [ObjectType] ,[Entity] , [EmailBody], [EmailSub],  [EmailID] , [Status], [sUser],[Seq],[DraftKey] ) " & _
                                    " VALUES('PR' , '1470000113', '" & p_oDICompany.CompanyName & "', '" & Replace(sBody, "'", "''") & "', '" & sEmailSubject & "' ,'" & odr("Email").ToString.Trim() & "', 'Pending', " & _
                                    "'/" & Left(odr("User").ToString.Trim(), 28) & "/'," & icount & ", '" & sDraftkey & "' ) "
                                        End If
                                    Next

                                    ' ''sInsertSQL = "INSERT INTO " & p_sHoldingEntity & ".. [AB_EmailStatus]  " & _
                                    ' ''  "([DocType] , [ObjectType] ,[Entity] ,  [EmailBody], [EmailSub],[EmailID] , [Status] ) " & _
                                    ' ''  " VALUES('PR' , '1470000113', '" & p_oDICompany.CompanyName & "', '" & Replace(sBody, "'", "''") & "', '" & sEmailSubject & "' ,'" & p_sEmailAddress & "', 'Open' ) "

                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Email Table Insert SQL " & sInsertSQL, sFuncName)
                                    oRset.DoQuery(sInsertSQL)
                                Else
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Authorizer Email address is Empty ", sFuncName)
                                End If
                                Approval = False
                            End If

                            If oForm.Title = "Purchase Request - Draft [Approved]" Then

                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("oForm.Title = Purchase Request - Draft [Approved]", sFuncName)
                                sCreator = oForm.Items.Item("U_AB_POCREATOR").Specific.string
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creator " & sCreator, sFuncName)
                                sSQLstring = "SELECT T0.[Email] FROM OSLP T0 WHERE T0.[SlpName] = '" & sCreator & "'"
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Fetching Email Id for Purchasing Unit " & sSQLstring, sFuncName)
                                oRset.DoQuery(sSQLstring)
                                sCreator = oRset.Fields.Item("Email").Value

                                sDraftkey = oForm.DataSources.DBDataSources.Item(0).GetValue("DocEntry", 0)
                                sComments = oForm.DataSources.DBDataSources.Item(0).GetValue("Comments", 0)
                                sPRNo = oForm.DataSources.DBDataSources.Item(0).GetValue("DocNum", 0)

                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting Document Total  " & sSQLstring, sFuncName)
                                sSQLstring = "SELECT Currency , cast( cast(sum(T0.[LineTotal]) as decimal(18,2)) as nvarchar) [LC], cast(cast(sum(T0.[TotalFrgn]) as decimal(18,2)) as nvarchar) [FC] FROM PRQ1 T0 WHERE T0.[DocEntry] = '" & sDraftkey & "' group by  Currency"
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Fetching Email Id for Purchasing Unit " & sSQLstring, sFuncName)
                                oRset.DoQuery(sSQLstring)
                                dDoctatalFC = oRset.Fields.Item("FC").Value
                                dDoctotalLC = oRset.Fields.Item("LC").Value
                                sCurrency = oRset.Fields.Item("Currency").Value

                                sEmailSubject = "Approved PR No. " & sPRNo & "  " & p_oDICompany.CompanyName & " has been approved for your action "
                                p_SyncDateTime = Format(Now, "dddd") & ", " & Format(Now, "MMM") & " " & Format(Now, "dd") & ", " & Format(Now, "yyyy") & " " & Format(Now, "HH:mm:ss")
                                sBody = sBody & "<div align=left style='font-size:10.0pt;font-family:Arial'>"
                                sBody = sBody & " Dear Sir/Madam,<br /><br />"
                                sBody = sBody & p_SyncDateTime & " <br /><br />"
                                sBody = sBody & " " & "<B> An Approved PR is in SAP for your action </B><br /><br />"
                                sBody = sBody & " " & "<B> Approved PR No. : " & sPRNo & " </B>  <br />"
                                sBody = sBody & " " & " Document Amount before GST : SGD " & dDoctotalLC & " <br />"
                                If CDbl(dDoctatalFC) > 0 Then
                                    sBody = sBody & " " & " Document Amount before GST : " & sCurrency & " " & dDoctatalFC & " <br />"
                                End If
                                ''sBody = sBody & " " & " Originator      : " & p_oDICompany.UserName & " <br />"
                                sBody = sBody & " " & " Entity          : " & p_oDICompany.CompanyName & " <br />"
                                sBody = sBody & " " & " Remarks        : " & sComments
                                sBody = sBody & "<br /><br />"
                                sBody = sBody & " Please login to the above mentioned entity for more details. <br /><br />"
                                sBody = sBody & " Thank you. <br /><br />"
                                sBody = sBody & " Please do not reply to this email. <div/>"

                                sInsertSQL += "INSERT INTO " & p_sHoldingEntity & ".. [AB_EmailStatus]  " & _
                        "([DocType] , [ObjectType] ,[Entity] , [EmailBody], [EmailSub],  [EmailID] , [Status], [sUser],[Seq],[DraftKey] ) " & _
                        " VALUES('PR' , '1470000113', '" & p_oDICompany.CompanyName & "', '" & Replace(sBody, "'", "''") & "', '" & sEmailSubject & "' ,'" & sCreator & "', 'Open', " & _
                        "'/" & p_oDICompany.UserName & "/'," & 1 & ", '" & sDraftkey & "' ) "

                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Email Table Insert SQL PR Approved " & sInsertSQL, sFuncName)
                                oRset.DoQuery(sInsertSQL)
                            End If

                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

                        Catch ex As Exception
                            p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                        Finally
                            oRset = Nothing

                        End Try

                End Select

            ElseIf BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE Then

                'Select Case BusinessObjectInfo.FormTypeEx

                '    ''----- Commented BP Approval as per the NIK advise
                '    Case "134"
                '        Dim oRset As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                '        Dim sQuery As String = String.Empty
                '        Dim dAmount As Double = 0
                '        Dim sEmailsender As String = String.Empty
                '        Dim sInsertSQL As String = String.Empty
                '        Dim sEmailSubject As String = String.Empty
                '        Dim sBody As String = String.Empty
                '        Dim p_SyncDateTime As String = String.Empty
                '        Dim sSQLstring As String = String.Empty
                '        Dim sDraftkey As String = String.Empty
                '        Dim sBPCode As String = String.Empty
                '        Dim sBPName As String = String.Empty
                '        Try
                '            oForm = p_oSBOApplication.Forms.GetFormByTypeAndCount(BusinessObjectInfo.FormTypeEx, p_FormTypecount)
                '            Dim oForm_UDF As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount(-134, p_FormTypecount)
                '            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("BP FormDataEvent - UPDATE Mode ", sFuncName)
                '            If oForm_UDF.Items.Item("U_AB_STATUS").Specific.value.ToString.Trim() = "PENDING" Then
                '                sBPCode = oForm.DataSources.DBDataSources.Item(0).GetValue("CardCode", 0).ToString.Trim()
                '                sBPName = oForm.DataSources.DBDataSources.Item(0).GetValue("CardName", 0).ToString.Trim()

                '                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("status is Pending ", sFuncName)

                '                sEmailSubject = "Request for BP approval in SAP - " & p_oDICompany.CompanyName & "  BP Code. " & sBPCode
                '                p_SyncDateTime = Format(Now, "dddd") & ", " & Format(Now, "MMM") & " " & Format(Now, "dd") & ", " & Format(Now, "yyyy") & " " & Format(Now, "HH:mm:ss")
                '                sBody = sBody & "<div align=left style='font-size:10.0pt;font-family:Arial'>"
                '                sBody = sBody & " Dear Sir/Madam,<br /><br />"
                '                sBody = sBody & p_SyncDateTime & " <br /><br />"
                '                sBody = sBody & " " & " <B> Request for your Business Partner approval in SAP . </B><br /><br />"
                '                sBody = sBody & " " & "<B> BP Code      : " & sBPCode & " </B> (Can be viewed under Main Menu/ Business Partners) <br />"
                '                sBody = sBody & " " & " BP name         : " & sBPName & " <br />"
                '                sBody = sBody & " " & " Originator      : " & p_oDICompany.UserName & " <br />"
                '                sBody = sBody & " " & " Entity          : " & p_oDICompany.CompanyName

                '                sBody = sBody & "<br /><br />"
                '                sBody = sBody & " Please login to the above mentioned entity to approve the document."
                '                sBody = sBody & "<br /><br />"
                '                sBody = sBody & " Thank you."
                '                sBody = sBody & "<br /><br />"
                '                sBody = sBody & " Please do not reply to this email. <div/>"

                '                sInsertSQL += "INSERT INTO " & p_sHoldingEntity & ".. [AB_EmailStatus]  " & _
                '        "([DocType] , [ObjectType] ,[Entity] , [EmailBody], [EmailSub],  [EmailID] , [Status], [sUser],[Seq],[DraftKey] ) " & _
                '        " VALUES('BP' , '2', '" & p_oDICompany.CompanyName & "', '" & Replace(sBody, "'", "''") & "', '" & sEmailSubject & "' ,'" & p_oCompDef.sApprover.Trim() & "', 'Open', " & _
                '        "'" & p_oDICompany.UserName & "'," & 1 & ", '0' ) "

                '                If sInsertSQL.Length > 0 Then
                '                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Insert in EmailStatus Table " & sInsertSQL, sFuncName)
                '                    oRset.DoQuery(sInsertSQL)
                '                End If

                '            End If


                '        Catch ex As Exception
                '            p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                '            Call WriteToLogFile(sErrDesc, sFuncName)
                '            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                '        Finally
                '            oRset = Nothing
                '        End Try
                'End Select
            End If
        End If
    End Sub


    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent

        Try
            If pVal.BeforeAction = True Then
                Select Case pVal.MenuUID

                    ''----------------Duplication menu trigger

                    '' ''Case "1287"
                    '' ''    Dim oform As SAPbouiCOM.Form = Nothing
                    '' ''    oform = p_oSBOApplication.Forms.ActiveForm

                    '' ''    If oform.TypeEx = "142" Then
                    '' ''        Dim oform_UDF As SAPbouiCOM.Form = Nothing
                    '' ''        oform_UDF = p_oSBOApplication.Forms.GetFormByTypeAndCount(-142, oform.TypeCount)
                    '' ''        Dim oMAtrix As SAPbouiCOM.Matrix = Nothing
                    '' ''        oMAtrix = oform.Items.Item("38").Specific
                    '' ''        Try
                    '' ''            oform_UDF.Items.Item("U_AB_WAIVER").Specific.String = String.Empty
                    '' ''            oform_UDF.Items.Item("U_AB_APPROVALAMT").Specific.String = String.Empty
                    '' ''        Catch ex As Exception
                    '' ''            p_oSBOApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    '' ''        End Try

                    '' ''    End If


                    'Case "1282"

                    '    '--- BP Approval uncomment

                    '    Dim oform As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                    '    Dim oform_UDF As SAPbouiCOM.Form = Nothing
                    '    If oform.TypeEx = "134" Then
                    '        Try
                    '            oform_UDF = p_oSBOApplication.Forms.GetFormByTypeAndCount(-134, oform.TypeCount)
                    '            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Menu Event " & p_oCompDef.sAuthorization, sFuncName)
                    '            If p_oCompDef.sAuthorization.Trim() = "CREATE AND UPDATE" Then
                    '                oform_UDF.Items.Item("U_AB_STATUS").Specific.select("PENDING")
                    '                oform.Items.Item("btnapprove").Enabled = False
                    '            End If
                    '        Catch ex As Exception
                    '            p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    '            Call WriteToLogFile(sErrDesc, sFuncName)
                    '            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                    '            Exit Sub
                    '        End Try

                    '        ''ElseIf oform.TypeEx = "142" Or oform.TypeEx = "1470000200" Then
                    '        ''    Try
                    '        ''        oform.Items.Item("txtnote").Enabled = False
                    '        ''    Catch ex As Exception
                    '        ''        p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    '        ''        Call WriteToLogFile(sErrDesc, sFuncName)
                    '        ''        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                    '        ''        Exit Sub
                    '        ''    End Try


                    '    End If
                    '--- Creating the Approve button

                    'Case "1289", "1290", "1291", "1288"
                    '    Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                    '    Try

                    '        Dim oform_UDF As SAPbouiCOM.Form = Nothing
                    '        If oForm.TypeEx = "134" Or oForm.TypeEx = "-134" Then
                    '            oform_UDF = p_oSBOApplication.Forms.GetFormByTypeAndCount(-134, oForm.TypeCount)
                    '            oForm = p_oSBOApplication.Forms.GetFormByTypeAndCount(134, oForm.TypeCount)
                    '            If p_oCompDef.sAuthorization = "APPROVE" And oform_UDF.Items.Item("U_AB_STATUS").Specific.value.ToString.Trim() <> "APPROVED" Then
                    '                oForm.Items.Item("btnapprove").Enabled = True
                    '            Else
                    '                oForm.Items.Item("btnapprove").Enabled = False
                    '            End If
                    '            ''ElseIf oForm.TypeEx = "142" Or oForm.TypeEx = "1470000200" Then
                    '            ''    Try
                    '            ''        oForm.Items.Item("txtnote").Enabled = False
                    '            ''    Catch ex As Exception
                    '            ''        p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    '            ''        Call WriteToLogFile(sErrDesc, sFuncName)
                    '            ''        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                    '            ''        Exit Sub
                    '            ''    End Try
                    '        End If
                    '    Catch ex As Exception
                    '        p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    '        Call WriteToLogFile(sErrDesc, sFuncName)
                    '        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                    '        Exit Sub
                    '    End Try

                    ''Case "1281"
                    ''    Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                    ''    If oForm.TypeEx = "142" Or oForm.TypeEx = "1470000200" Then
                    ''        Try
                    ''            oForm.Items.Item("txtnote").Enabled = False
                    ''        Catch ex As Exception
                    ''            p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    ''            Call WriteToLogFile(sErrDesc, sFuncName)
                    ''            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                    ''            Exit Sub
                    ''        End Try
                    ''    End If
                    Case "1283", "1284", "1286"

                        Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                        If oForm.TypeEx = "3002" Or oForm.TypeEx = "50105" Then
                            Try
                                If p_sAppStatus = "W" Then
                                    p_oSBOApplication.StatusBar.SetText("Can`t Cancel / Close / Remove the Document which is triggered for an approval ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    BubbleEvent = False
                                    Exit Sub
                                End If

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

Option Explicit On
Imports SAPbouiCOM.Framework
Imports System.Windows.Forms
Imports System.Text.RegularExpressions


Namespace AE_PWC_AO02
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
            Dim oCombo As SAPbouiCOM.ComboBox = Nothing
            Try
                sFuncName = "SBO_Application_MenuEvent()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

                If Not p_oDICompany.Connected Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectDICompSSO()", sFuncName)
                    If ConnectDICompSSO(p_oDICompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                End If

                If pVal.BeforeAction = False Then
                    Select Case pVal.MenuUID

                        Case "FGTF"
                            Try
                                LoadFromXML("GenerateTextFile.srf", SBO_Application)
                                oForm = p_oSBOApplication.Forms.Item("GTF")
                                oForm.Visible = True
                                If EntityLoad(oForm, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                                Exit Try
                            Catch ex As Exception
                                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                            End Try
                            Exit Sub
                        Case "BCA"
                            Try
                                LoadFromXML("CostAllocation.srf", SBO_Application)
                                oForm = p_oSBOApplication.Forms.Item("CA")
                                oForm.Freeze(True)
                                oCombo = oForm.Items.Item("Item_13").Specific
                                oCombo.ValidValues.Add("Jan", "1")
                                oCombo.ValidValues.Add("Feb", "2")
                                oCombo.ValidValues.Add("Mar", "3")
                                oCombo.ValidValues.Add("Apr", "4")
                                oCombo.ValidValues.Add("May", "5")
                                oCombo.ValidValues.Add("Jun", "6")
                                oCombo.ValidValues.Add("Jul", "7")
                                oCombo.ValidValues.Add("Aug", "8")
                                oCombo.ValidValues.Add("Sep", "9")
                                oCombo.ValidValues.Add("Oct", "10")
                                oCombo.ValidValues.Add("Nov", "11")
                                oCombo.ValidValues.Add("Dec", "12")
                                oCombo = oForm.Items.Item("Item_14").Specific
                                oCombo.ValidValues.Add("Jan", "1")
                                oCombo.ValidValues.Add("Feb", "2")
                                oCombo.ValidValues.Add("Mar", "3")
                                oCombo.ValidValues.Add("Apr", "4")
                                oCombo.ValidValues.Add("May", "5")
                                oCombo.ValidValues.Add("Jun", "6")
                                oCombo.ValidValues.Add("Jul", "7")
                                oCombo.ValidValues.Add("Aug", "8")
                                oCombo.ValidValues.Add("Sep", "9")
                                oCombo.ValidValues.Add("Oct", "10")
                                oCombo.ValidValues.Add("Nov", "11")
                                oCombo.ValidValues.Add("Dec", "12")
                                oForm.Items.Item("Item_17").Width = 250
                                oForm.Items.Item("Item_17").Height = 15
                                oForm.Items.Item("Item_18").Width = 200
                                oForm.Items.Item("Item_18").Height = 15

                                oForm.Freeze(False)
                                oForm.Visible = True
                                oForm.Items.Item("Item_17").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                oForm.PaneLevel = 2

                                Exit Try
                            Catch ex As Exception
                                oForm.Freeze(False)
                                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                            End Try
                            Exit Sub

                        Case "1284"
                            oForm = p_oSBOApplication.Forms.ActiveForm
                            If oForm.TypeEx = "392" Then
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Set bJEflag = Yes", sFuncName)
                                bJEflag = True
                            End If
                    End Select
                End If

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
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
            Dim p_oDVJE As DataView = Nothing
            Dim oDTDistinct As DataTable = Nothing
            Dim oDTRowFilter As DataTable = Nothing

            Try
                sFuncName = "SBO_Application_ItemEvent()"
                ' If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

                If Not IsNothing(p_oDICompany) Then
                    If Not p_oDICompany.Connected Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectDICompSSO()", sFuncName)
                        If ConnectDICompSSO(p_oDICompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    End If
                End If

                If pVal.BeforeAction = False Then

                    Select Case pVal.FormUID

                        Case "GTF"
                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                                If pVal.ItemUID = "btnBrowse" Then
                                    Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                                    Try
                                        sFuncName = "'Browse' Button Click - ID 'btnBrowse'"
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling File Open Function", sFuncName)

                                        fillopen()

                                        oForm.Items.Item("txtFldPath").Specific.string = p_sSelectedFilepath
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With Success File Open Function", sFuncName)
                                    Catch ex As Exception
                                        BubbleEvent = False
                                        sErrDesc = ex.Message
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                                        WriteToLogFile(Err.Description, sFuncName)
                                        ShowErr(sErrDesc)
                                    End Try

                                    ' oForm.Items.Item("Item_5").Specific.string = p_sSelectedFilepath
                                    Exit Sub
                                End If
                            End If

                        Case "CA"

                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                oCFLEvento = pVal
                                Dim sCFL_ID As String
                                sCFL_ID = oCFLEvento.ChooseFromListUID
                                Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.Item(FormUID)
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                oCFL = oForm.ChooseFromLists.Item(sCFL_ID)
                                Dim omatrix As SAPbouiCOM.Matrix = Nothing

                                Try
                                    If oCFLEvento.BeforeAction = False Then
                                        Dim oDataTable As SAPbouiCOM.DataTable
                                        oDataTable = oCFLEvento.SelectedObjects
                                        If pVal.ItemUID = "Item_10" Then 'Acct Code
                                            oForm.Items.Item("Item_10").Specific.string = oDataTable.GetValue("AcctCode", 0)
                                        End If
                                        If pVal.ItemUID = "Item_12" Then 'Acct Code
                                            oForm.Items.Item("Item_12").Specific.string = oDataTable.GetValue("AcctCode", 0)
                                        End If
                                        If pVal.ItemUID = "Item_5" Then 'Acct Code
                                            oForm.Items.Item("Item_5").Specific.string = oDataTable.GetValue("OcrCode", 0)
                                        End If
                                        If pVal.ItemUID = "Item_8" Then 'Acct Code
                                            oForm.Items.Item("Item_8").Specific.string = oDataTable.GetValue("OcrCode", 0)
                                        End If
                                        If pVal.ItemUID = "Item_24" And pVal.ColUID = "Col_6" Then 'New OU
                                            omatrix = oForm.Items.Item("Item_24").Specific
                                            omatrix.Columns.Item("Col_6").Cells.Item(pVal.Row).Specific.String = oDataTable.GetValue("OcrCode", 0)
                                        End If

                                        If pVal.ItemUID = "Item_24" And pVal.ColUID = "V_5" Then ' GL Account
                                            omatrix = oForm.Items.Item("Item_24").Specific
                                            omatrix.Columns.Item("V_5").Cells.Item(pVal.Row).Specific.String = oDataTable.GetValue("AcctCode", 0)
                                        End If

                                        If pVal.ItemUID = "Item_2" Then 'Acct Code
                                            oForm.Items.Item("Item_2").Specific.string = oDataTable.GetValue("PrcCode", 0)
                                        End If
                                        If pVal.ItemUID = "Item_15" Then 'Acct Code
                                            oForm.Items.Item("Item_15").Specific.string = oDataTable.GetValue("PrcCode", 0)
                                        End If

                                        If pVal.ItemUID = "Item_22" Then 'Acct Code
                                            oForm.Items.Item("Item_22").Specific.string = oDataTable.GetValue("PrcCode", 0)
                                        End If
                                        If pVal.ItemUID = "Item_26" Then 'Acct Code
                                            oForm.Items.Item("Item_26").Specific.string = oDataTable.GetValue("PrcCode", 0)
                                        End If

                                    End If
                                Catch ex As Exception
                                End Try
                            End If


                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                                If pVal.ItemUID = "Item_17" Or pVal.ItemUID = "Item_18" Then
                                    Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                                    Select Case pVal.ItemUID
                                        Case "Item_17"
                                            oForm.PaneLevel = 2
                                        Case "Item_18"
                                            oForm.PaneLevel = 3
                                    End Select
                                End If
                            End If

                        Case "CostCentres"
                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE Then
                                Dim oform As SAPbouiCOM.Form = p_oSBOApplication.Forms.Item("CostCentres")
                                oform.Visible = True
                                BubbleEvent = False
                                Exit Sub
                            End If
                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE Then
                                p_oCostCentre.Clear()
                            End If

                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                                Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                                If pVal.ItemUID = "Item_2" And pVal.ColUID = "Col_0" Then
                                    Dim oCheck As SAPbouiCOM.CheckBox
                                    Dim oMatrix As SAPbouiCOM.Matrix = Nothing
                                    Dim sValue As String = String.Empty

                                    If p_bCostcentre = False Then
                                        oMatrix = oForm.Items.Item("Item_2").Specific
                                        oCheck = oMatrix.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific
                                        sValue = oMatrix.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific.String
                                        If oCheck.Checked = True Then
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Added " & sValue, sFuncName)
                                            p_oCostCentre.Rows.Add(sValue)
                                        Else
                                            Dim foundRow() As DataRow = p_oCostCentre.Select("OU='" & sValue & "'")
                                            If foundRow.Count > 0 Then
                                                For Each row As DataRow In foundRow
                                                    row.Delete()
                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Deleted " & sValue, sFuncName)
                                                Next
                                            End If
                                        End If
                                    Else
                                        p_bCostcentre = False
                                    End If


                                ElseIf pVal.ItemUID = "21" Then
                                    p_Summaryreport = String.Empty
                                    If p_oCostCentre.Rows.Count > 0 Then
                                        For Each odr As DataRow In p_oCostCentre.Rows
                                            p_Summaryreport += "'" & odr(0) & "',"
                                        Next
                                        p_Summaryreport = Left(p_Summaryreport, p_Summaryreport.Length - 1)
                                        oForm.Close()
                                    Else
                                        p_oSBOApplication.StatusBar.SetText("Operating Unit should not be blank .......!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                            End If

                        Case "Distributionrules"
                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE Then
                                Dim oform As SAPbouiCOM.Form = p_oSBOApplication.Forms.Item("Distributionrules")
                                oform.Visible = True
                                BubbleEvent = False
                                Exit Sub
                            End If
                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE Then
                                p_oDistribution.Clear()
                            End If

                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                                Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                                If pVal.ItemUID = "Item_2" And pVal.ColUID = "Col_0" Then
                                    Dim oCheck As SAPbouiCOM.CheckBox
                                    Dim oMatrix As SAPbouiCOM.Matrix = Nothing
                                    Dim sValue As String = String.Empty
                                    If p_bDimension = False Then
                                        oMatrix = oForm.Items.Item("Item_2").Specific
                                        oCheck = oMatrix.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific
                                        sValue = oMatrix.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific.String
                                        If oCheck.Checked = True Then
                                            p_oDistribution.Rows.Add(sValue)
                                        Else
                                            Dim foundRow() As DataRow = p_oDistribution.Select("OU='" & sValue & "'")
                                            If foundRow.Count > 0 Then
                                                For Each row As DataRow In foundRow
                                                    row.Delete()
                                                Next
                                            End If
                                        End If
                                    Else
                                        p_bDimension = False
                                    End If


                                ElseIf pVal.ItemUID = "21" Then
                                    p_Dimensionrules = String.Empty
                                    If p_oDistribution.Rows.Count > 0 Then
                                        For Each odr As DataRow In p_oDistribution.Rows
                                            p_Dimensionrules += "'" & odr(0) & "',"
                                        Next
                                        p_Dimensionrules = Left(p_Dimensionrules, p_Dimensionrules.Length - 1)
                                        oForm.Close()
                                    Else
                                        p_oSBOApplication.StatusBar.SetText("Dimension rules should not be blank .......!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                            End If


                    End Select
                Else
                    Select Case pVal.FormTypeEx
                        Case "392"
                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                                If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And bJEflag = True Then
                                    Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetFormByTypeAndCount(pVal.FormTypeEx, pVal.FormTypeCount)
                                    Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("76").Specific
                                    p_sCancel = String.Empty
                                    For imjs As Integer = 1 To oMatrix.RowCount
                                        If Not String.IsNullOrEmpty(oMatrix.Columns.Item("2003").Cells.Item(imjs).Specific.String) And _
                                            Not String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_JV").Cells.Item(imjs).Specific.String) And _
                                            Not String.IsNullOrEmpty(oMatrix.Columns.Item("U_AB_OcrCode3").Cells.Item(imjs).Specific.String) Then
                                            p_sCancel += " delete from [@AB_COSTALLOCATION] where U_Transid = '" & oMatrix.Columns.Item("U_AB_JV").Cells.Item(imjs).Specific.String & "' " & _
                                                "and U_SourceOcrcode3 = '" & oMatrix.Columns.Item("2003").Cells.Item(imjs).Specific.String & "' " & _
                                                "and U_OcrCode3 = '" & oMatrix.Columns.Item("U_AB_OcrCode3").Cells.Item(imjs).Specific.String & "' "
                                        End If
                                    Next
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Cancelling JE " & p_sCancel, sFuncName)
                                End If
                            End If
                    End Select
                    Select Case pVal.FormUID

                        Case "CA"

                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE Then
                                p_Dimensionrules = String.Empty
                            End If

                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then

                                If pVal.ItemUID = "30" Then
                                    Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                                    Dim sSQL As String = String.Empty
                                    Dim oCheck As SAPbouiCOM.CheckBox = Nothing
                                    Try
                                        LoadFromXML("CFL_Costcenter.srf", SBO_Application)
                                        oForm = p_oSBOApplication.Forms.Item("CostCentres")
                                        oForm.Freeze(True)
                                        p_bCostcentre = True
                                        Try
                                            oForm.DataSources.DataTables.Add("OPRC")
                                        Catch ex As Exception
                                        End Try
                                        Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("Item_2").Specific

                                        sSQL = "SELECT T0.[PrcCode], T0.[PrcName], 'N' [Active], U_AB_ENTITYNAME FROM OPRC T0  WHERE T0.[DimCode] = 3 order by U_AB_ENTITYNAME"
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("CFL Load " & sSQL, sFuncName)
                                        Dim oSApDT As SAPbouiCOM.DataTable
                                        oSApDT = oForm.DataSources.DataTables.Item("OPRC")
                                        oSApDT.ExecuteQuery(sSQL)
                                        '' oMatrix.Clear()
                                        p_oCostCentre = New DataTable
                                        p_oCostCentre.Columns.Add("OU", GetType(String))

                                        oForm.Items.Item("Item_2").Specific.columns.item("Col_0").databind.bind("OPRC", "Active")
                                        oForm.Items.Item("Item_2").Specific.columns.item("Col_1").databind.bind("OPRC", "PrcCode")
                                        oForm.Items.Item("Item_2").Specific.columns.item("Col_2").databind.bind("OPRC", "PrcName")
                                        oForm.Items.Item("Item_2").Specific.columns.item("Col_3").databind.bind("OPRC", "U_AB_ENTITYNAME")
                                        oForm.Items.Item("Item_2").Specific.LoadFromDataSource()
                                        oForm.Items.Item("Item_2").Specific.AutoResizeColumns()
                                        oCheck = oMatrix.Columns.Item("Col_0").Cells.Item(1).Specific
                                        oCheck.Checked = True
                                        oCheck.Checked = False
                                        oForm.Freeze(False)
                                        oForm.Visible = True
                                    Catch ex As Exception
                                        p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        oForm.Freeze(False)
                                        BubbleEvent = False
                                        Exit Sub
                                    End Try
                                End If

                                If pVal.ItemUID = "1000002" Then
                                    Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                                    Dim sSQL As String = String.Empty
                                    Dim oCheck As SAPbouiCOM.CheckBox = Nothing
                                    Try
                                        LoadFromXML("CFL_Distribution.srf", SBO_Application)
                                        oForm = p_oSBOApplication.Forms.Item("Distributionrules")
                                        oForm.Freeze(True)
                                        p_bDimension = True
                                        Try
                                            oForm.DataSources.DataTables.Add("OOCR")
                                        Catch ex As Exception
                                        End Try
                                        Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("Item_2").Specific

                                        sSQL = "SELECT T0.[OcrCode], T0.[OcrName], 'N' [Active] FROM OOCR T0 WHERE T0.[DimCode] = 3"
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("CFL Load " & sSQL, sFuncName)
                                        Dim oSApDT As SAPbouiCOM.DataTable
                                        oSApDT = oForm.DataSources.DataTables.Item("OOCR")
                                        oSApDT.ExecuteQuery(sSQL)
                                        '' oMatrix.Clear()
                                        p_oDistribution = New DataTable
                                        p_oDistribution.Columns.Add("OU", GetType(String))

                                        oForm.Items.Item("Item_2").Specific.columns.item("Col_0").databind.bind("OOCR", "Active")
                                        oForm.Items.Item("Item_2").Specific.columns.item("Col_1").databind.bind("OOCR", "OcrCode")
                                        oForm.Items.Item("Item_2").Specific.columns.item("Col_2").databind.bind("OOCR", "OcrName")
                                        oForm.Items.Item("Item_2").Specific.LoadFromDataSource()
                                        oForm.Items.Item("Item_2").Specific.AutoResizeColumns()
                                        oCheck = oMatrix.Columns.Item("Col_0").Cells.Item(1).Specific
                                        oCheck.Checked = True
                                        oCheck.Checked = False
                                        oForm.Freeze(False)
                                        oForm.Visible = True
                                    Catch ex As Exception
                                        p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        oForm.Freeze(False)
                                        BubbleEvent = False
                                        Exit Sub
                                    End Try
                                End If

                                If pVal.ItemUID = "Item_19" Then
                                    Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                                    Dim sMonth As String = String.Empty
                                    Dim sSplitM() As String
                                    Dim sSplitD() As String
                                    Dim sSplitG() As String
                                    Dim sDistribution As String = String.Empty
                                    Dim sGLAccount As String = String.Empty
                                    Dim sSQL As String = String.Empty
                                    Dim oGridT1 As SAPbouiCOM.Grid = Nothing
                                    Dim oGridT2 As SAPbouiCOM.Grid = Nothing
                                    Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("Item_24").Specific
                                    Dim sNow As String = String.Empty
                                    Dim sIN As String = String.Empty
                                    Dim sQuotes As String = String.Empty
                                    Dim oRset As SAPbobsCOM.Recordset = Nothing

                                    Try
                                        oRset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        sFuncName = "Header Show button click()"
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling HeaderValidation()", sFuncName)
                                        oForm.Items.Item("Item_17").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                                        p_oSBOApplication.SetStatusBarMessage("Cost Allocation Information Starting to Load .......!", SAPbouiCOM.BoMessageTime.bmt_Short, False)

                                        If CostAllocation_Validation(oForm, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                        oForm.Freeze(True)
                                        sMonth = oForm.Items.Item("Item_13").Specific.selected.description & "," & oForm.Items.Item("Item_14").Specific.selected.description
                                        sDistribution = oForm.Items.Item("Item_5").Specific.value & "," & oForm.Items.Item("Item_8").Specific.value
                                        sGLAccount = oForm.Items.Item("Item_10").Specific.value & "," & oForm.Items.Item("Item_12").Specific.value
                                        sSplitM = sMonth.Split(",")
                                        '' sDistribution = p_Dimensionrules
                                        sSplitG = sGLAccount.Split(",")
                                        ''OcrCode3 U_AB_NONPROJECT
                                        '' sSQL = "SELECT T0.[DocNum] , T0.[DocType] FROM OPCH T0 union all SELECT '' [DocNum] , '' [DocType] "

                                        'sSQL = "DECLARE  " & _
                                        '   "  @string varchar(100), @string1 varchar(100), @string2 varchar(100), @string3 varchar(max), @string4 varchar(max) " & _
                                        '   "  SET @string = '" & sSplitD(0) & "' SET @string1 = '" & sSplitD(1) & "' SET @string2 = '" & sSplitD(0) & "' " & _
                                        '   " WHILE PATINDEX('%[^a-z]%',@string2) > 0 SET @string2 = STUFF(@string2,PATINDEX('%[^a-z]%',@string2),1,'') " & _
                                        '   " WHILE PATINDEX('%[^0-9]%',@string) <> 0     SET @string = STUFF(@string,PATINDEX('%[^0-9]%',@string),1,'') " & _
                                        '   " WHILE PATINDEX('%[^0-9]%',@string1) <> 0     SET @string1 = STUFF(@string1,PATINDEX('%[^0-9]%',@string1),1,'') " & _
                                        '   " WHILE cast(@string as numeric ) <= cast(@string1 as numeric ) " & _
                                        '   " begin " & _
                                        '   " SET @string3 =  isnull(@string3,'')  + '[' + @string2 + @string + '],' " & _
                                        '   " set @string = cast(@string as numeric ) + 1 " & _
                                        '   " end " & _
                                        '   " set @string3 = replace( replace(@string3,'[',''''),']','''') " & _
                                        '   " set @string3 = left(@string3, len(@string3) -1) " & _
                                        '     " set @string4 = replace(@string3,'''','''''') " & _
                                        '   " select @string3 [Ouput] , @string4 [quotes] "
                                        'oRset.DoQuery(sSQL)
                                        'sIN = oRset.Fields.Item("Ouput").Value
                                        'sQuotes = oRset.Fields.Item("quotes").Value

                                        sIN = p_Dimensionrules
                                        sQuotes = sIN.Replace("'", "''")

                                        sSQL = "Select A.DOCNUM 'SAP Reference Number' , A.NUMATCARD 'Vendor Invoice Number',A.TAXDATE 'Vendor Invoice Date',A.DOCDATE 'SAP Posting Date',B.ACCTCODE 'Gl Accounts',C.ACCTNAME 'Gl Description',A.CARDNAME 'Vendor Name',A.JRNLMEMO 'Journal Remark', B.LineTotal 'Invoice Amount', B.[OcrCode3] 'Distribution Code',D.SLPNAME 'Purchasing Department'" & _
                                        " From OPCH A INNER JOIN PCH1 B ON A.DOCENTRY=B.DOCENTRY LEFT JOIN OACT C ON B.ACCTCODE=C.ACCTCODE LEFT JOIN OSLP D ON A.SLPCODE=D.SLPCODE" & _
                                        " WHERE month(A.DOCDATE)>= " & sSplitM(0) & " and month(A.docdate)<= " & sSplitM(0) & " and year(A.DOCDATE) = " & Now.Year & "" & _
                                        " and B.[OcrCode3] in (" & sIN & ") and isnull(A.Indicator,'') <> 'CA' and c.GroupMask = 5 " & _
                                        " and B.[AcctCode] >= '" & sSplitG(0) & "' and B.[AcctCode] <= '" & sSplitG(1) & "'" & _
                                        " union all " & _
                                        "Select A.DOCNUM 'SAP Reference Number' , A.NUMATCARD 'Vendor Invoice Number',A.TAXDATE 'Vendor Invoice Date',A.DOCDATE 'SAP Posting Date',B.ACCTCODE 'Gl Accounts',C.ACCTNAME 'Gl Description',A.CARDNAME 'Vendor Name',A.JRNLMEMO 'Journal Remark', -1 * B.LineTotal 'Invoice Amount', B.[OcrCode3] 'Distribution Code',D.SLPNAME 'Purchasing Department'" & _
                                        " From ORPC A INNER JOIN RPC1 B ON A.DOCENTRY=B.DOCENTRY LEFT JOIN OACT C ON B.ACCTCODE=C.ACCTCODE LEFT JOIN OSLP D ON A.SLPCODE=D.SLPCODE" & _
                                        " WHERE month(A.DOCDATE)>= " & sSplitM(0) & " and month(A.docdate)<= " & sSplitM(0) & " and year(A.DOCDATE) = " & Now.Year & "" & _
                                        " and B.[OcrCode3] in (" & sIN & ")  and isnull(A.Indicator,'') <> 'CA' and c.GroupMask = 5 " & _
                                        " and B.[AcctCode] >= '" & sSplitG(0) & "' and B.[AcctCode] <= '" & sSplitG(1) & "'" & _
                                        " union all " & _
                                        "Select A.Number 'SAP Reference Number',B.REF2 'Vendor Invoice Number',B.TAXDATE 'Vendor Invoice Date', A.REFDATE 'SAP Posting Date', B.Account 'Gl Accounts', C.ACCTNAME 'Gl Description', 'JE' 'Vendor Name',A.MEMO 'Journal Remark', b.Debit - b.Credit 'Invoice Amount', B.[OcrCode3] 'Distribution Code','Purchasing Department' " & _
                                        "from ojdt A INNER JOIN JDT1 B ON A.TRANSID=B.TRANSID LEFT JOIN OACT C ON B.ACCOUNT=C.ACCTCODE " & _
                                        " where  a.transType =('30') and month(A.refdate)>= " & sSplitM(0) & " and month(A.refdate)<=" & sSplitM(0) & " and year(A.refdate) = " & Now.Year & "" & _
                                          " and B.[OcrCode3] in (" & sIN & ")  and isnull(A.Indicator,'') <> 'CA' and c.GroupMask = 5 " & _
                                        " and B.[Account] >= '" & sSplitG(0) & "' and B.[Account] <= '" & sSplitG(1) & "'"

                                        ''If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Tab 1 " & sSQL, sFuncName)
                                        ''oGridT1 = oForm.Items.Item("Item_22").Specific
                                        oGridT2 = oForm.Items.Item("Item_23").Specific
                                        ''Try
                                        ''    oForm.DataSources.DataTables.Add("MyDataTable")
                                        ''Catch ex As Exception
                                        ''End Try
                                        ''oGridT1.DataTable = oForm.DataSources.DataTables.Item("MyDataTable")
                                        ''oForm.DataSources.DataTables.Item(0).ExecuteQuery(sSQL)
                                        ''oGridT1.CollapseLevel = 1
                                        ''oGridT1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
                                        ''oGridT1.AutoResizeColumns()

                                        sNow = CStr(Now.Year) & sSplitM(0).PadLeft(2, "0"c) & "01"


                                        sSQL = "DECLARE @cols AS NVARCHAR(MAX),    @query  AS VARCHAR(max), @query1  AS VARCHAR(max), @cols1 as nvarchar(max) " & _
                                              "select @cols = STUFF((SELECT distinct  ',' + QUOTENAME(cast(T1.PrcCode as nvarchar(100))) " & _
                   " from OOCR T0  INNER JOIN OCR1 T1 ON T0.[OcrCode] = T1.[OcrCode]  where T0.OcrCode in (" & sIN & ") and T1.[ValidFrom] <= '" & sNow & "' and  (T1.[ValidTo] >= '" & sNow & "' or  isnull(T1.[ValidTo],'') = '')" & _
           " FOR XML PATH(''), TYPE " & _
           " ).value('.', 'NVARCHAR(MAX)') " & _
       "  ,1,1,'') " & _
    "select @cols1 = STUFF((SELECT '+ isnull(' + QUOTENAME(isnull(T1.PrcCode,0)) + ',0)' " & _
                   " from OOCR T0  INNER JOIN OCR1 T1 ON T0.[OcrCode] = T1.[OcrCode]  where T0.OcrCode in (" & sIN & ") and T1.[ValidFrom] <= '" & sNow & "' and  (T1.[ValidTo] >= '" & sNow & "' or  isnull(T1.[ValidTo],'') = '')" & _
           " FOR XML PATH(''), TYPE " & _
           " ).value('.', 'NVARCHAR(MAX)') " & _
       "  ,1,1,'') " & _
    " set @query = cast('SELECT SeriesName , ''PU '' + cast(DOCNUM as varchar(30)) ''SAP Reference Number'' , PDocNum ''Payment Doc Num'', PDocDate ''Payment Doc Date'' ,  NUMATCARD ''Vendor Invoice Number'',TAXDATE ''Vendor Invoice Date'',DOCDATE ''SAP Posting Date'',  " & _
    " CARDNAME ''Vendor / Payee Name'',JRNLMEMO ''Journal Remark'',  [OcrCode3] ''Distribution Code'', SLPNAME ''Purchasing Department'', Creator [Document Creator] , " & _
    "plan_id [Expenses], CAST(Line AS DECIMAL(19,3)) [Total Bill],' + @cols + '  " & _
                                          " from (" & _
               " SELECT i.SeriesName, f.DOCNUM   , f.NUMATCARD  ,f.TAXDATE  ,f.DOCDATE  ,f.CARDNAME  ,f.JRNLMEMO  ,  b.[OcrCode3]  ,h.SLPNAME  ," & _
               "a.AcctCode + '' - '' + a.AcctName [plan_id] , e.[PrcCode] [U_AB_NONPROJECT], l.DocNum [PDocNum] , l.DocDate [PDocDate] , g.LineNum , " & _
                "  case when g.BaseType = 18 then -sum(isnull(g.LineTotal,0)) else sum(isnull(g.LineTotal,0)) end [Line] , " & _
                " case when g.BaseType = 18 then -sum( round(isnull(g.LineTotal,0) /(d.[OcrTotal] /e.[PrcAmount] ),3)) else sum( round(isnull(g.LineTotal,0) /(d.[OcrTotal] /e.[PrcAmount] ),3)) end AS total , j.[U_NAME] +'', '' + j.[USER_CODE][Creator] " & _
                " FROM JDT1 b JOIN OACT a ON a.AcctCode = b.Account " & _
    " JOIN OJDT c on b.[TransId] = c.[TransId] " & _
    " left join OOCR d on d.OcrCode = b.OcrCode3 " & _
    " join OCR1 e on d.[OcrCode] = e.[OcrCode] " & _
    " JOIN OPCH f on c.[TransId] = f.[TransId]  JOIN PCH1 g on g.[DocEntry] = f.[DocEntry] and g.AcctCode = b.Account and  g.[OcrCode3] =  b.[OcrCode3] " & _
    "  LEFT JOIN OSLP h on f.SLPCODE=h.SLPCODE " & _
    "  LEFT JOIN NNM1 i on i.Series=f.Series " & _
    "   LEFT JOIN OUSR j on  f.[usersign] = j.[USERID] " & _
    " left join   [VPM2] k on f.DocEntry = k.DocEntry left JOIN OVPM l ON k.[DocNum] = l.[DocEntry] " & _
    " where  b.OcrCode3 in (" & sQuotes & ") and a.GroupMask = 5  and isnull(c.Indicator,'''') <> ''CA'' and " & _
    "	b.Account >= ''" & sSplitG(0) & "'' and b.Account <= ''" & sSplitG(1) & "'' and c.transType IN (''18'') and month(c.refdate) >= " & sSplitM(0) & " and month(c.refdate) <= " & sSplitM(1) & " and year(c.refdate) = " & Now.Year & " " & _
    " and  e.[ValidFrom] <= ''" & sNow & "'' and (e.[ValidTo] >= ''" & sNow & "'' or  isnull(e.[ValidTo],'''') = '''')" & _
     " and e.PrcCode not in (SELECT T0.[U_PrcCode] FROM [dbo].[@AB_COSTALLOCATION]  T0 WHERE T0.[U_Transid] = c.TransId and T0.U_OcrCode3 = b.OcrCode3) and (case when isnull(l.DocNum,'''') = '''' then ''N'' else l.Canceled end)=''N'' " & _
             "  group by a.AcctCode,a.AcctName,e.[PrcCode] , f.DOCNUM   , f.NUMATCARD  ,f.TAXDATE  ,f.DOCDATE  ,f.CARDNAME  ,f.JRNLMEMO  ,  b.[OcrCode3]  ,h.SLPNAME , g.LineNum, i.SeriesName ,j.[U_NAME] , j.[USER_CODE], l.DocNum , l.DocDate , g.BaseType  ) x pivot ( sum(total) " & _
             "   for U_AB_NONPROJECT in (' + @cols + ') " & _
           " ) p  " & _
               " union all SELECT SeriesName, ''PC '' + cast(DOCNUM as varchar(30)) ''SAP Reference Number'' , '''' ''Payment Doc Num'', NULL ''Payment Doc Date'',  NUMATCARD ''Vendor Invoice Number'',TAXDATE ''Vendor Invoice Date'',DOCDATE ''SAP Posting Date'',  " & _
    " CARDNAME ''Vendor / Payee Name'',JRNLMEMO ''Journal Remark'',  [OcrCode3] ''Distribution Code'', SLPNAME ''Purchasing Department'', Creator [Document Creator], " & _
    "plan_id [Expenses], Line [Total Bill],' + @cols + '  " & _
                                          " from (" & _
               " SELECT i.SeriesName, f.DOCNUM   , f.NUMATCARD  ,f.TAXDATE  ,f.DOCDATE  ,f.CARDNAME  ,f.JRNLMEMO  ,  b.[OcrCode3]  ,h.SLPNAME  , j.[U_NAME] +'', '' + j.[USER_CODE][Creator] ," & _
               "a.AcctCode + '' - '' + a.AcctName [plan_id] , e.[PrcCode] [U_AB_NONPROJECT], g.LineNum, " & _
                " case when g.BaseType = ''19'' then   sum(ISNULL( g.LineTotal,0))  else  - sum(ISNULL( g.LineTotal,0))  end [Line] , " & _
                " case when g.BaseType = ''19'' then  sum( round((isnull( g.LineTotal,0)) /(d.[OcrTotal] /e.[PrcAmount] ),3)) else - sum( round((isnull( g.LineTotal,0)) /(d.[OcrTotal] /e.[PrcAmount] ),3)) end AS total " & _
                " FROM JDT1 b JOIN OACT a ON a.AcctCode = b.Account " & _
    " JOIN OJDT c on b.[TransId] = c.[TransId] " & _
    " left join OOCR d on d.OcrCode = b.OcrCode3 " & _
    " join OCR1 e on d.[OcrCode] = e.[OcrCode] " & _
    " JOIN ORPC f on c.[TransId] = f.[TransId]  JOIN RPC1 g on g.[DocEntry] = f.[DocEntry] and g.AcctCode = b.Account  and  g.[OcrCode3] =  b.[OcrCode3] " & _
    "  LEFT JOIN OSLP h on f.SLPCODE=h.SLPCODE " & _
    "  LEFT JOIN NNM1 i on i.Series=f.Series " & _
    "   LEFT JOIN OUSR j on  f.[usersign] = j.[USERID] " & _
    " where  b.OcrCode3 in (" & sQuotes & ") and a.GroupMask = 5  and isnull(c.Indicator,'''') <> ''CA'' and " & _
    "	b.Account >= ''" & sSplitG(0) & "'' and b.Account <= ''" & sSplitG(1) & "'' and c.transType IN (''19'') and month(c.refdate) >= " & sSplitM(0) & " and month(c.refdate) <= " & sSplitM(1) & " and year(c.refdate) = " & Now.Year & " " & _
    " and  e.[ValidFrom] <= ''" & sNow & "'' and (e.[ValidTo] >= ''" & sNow & "'' or  isnull(e.[ValidTo],'''') = '''')" & _
     " and e.PrcCode not in (SELECT T0.[U_PrcCode] FROM [dbo].[@AB_COSTALLOCATION]  T0 WHERE T0.[U_Transid] = c.TransId and T0.U_OcrCode3 = b.OcrCode3) " & _
             "  group by a.AcctCode,a.AcctName,e.[PrcCode] , f.DOCNUM   , f.NUMATCARD  ,f.TAXDATE  ,f.DOCDATE  ,f.CARDNAME  ,f.JRNLMEMO  ,  b.[OcrCode3]  ,h.SLPNAME , g.LineNum , i.SeriesName ,j.[U_NAME] , j.[USER_CODE] ,g.BaseType ) x pivot ( sum(total) " & _
             "   for U_AB_NONPROJECT in (' + @cols + ') " & _
           " ) p ' as nvarchar(max)) " & _
  " set @query1 = cast( ' union all SELECT SeriesName, ''JE '' + cast(Number as varchar(30)) ''SAP Reference Number''  , '''' ''Payment Doc Num'', NULL ''Payment Doc Date'' ,REF2 ''Vendor Invoice Number'' , TAXDATE ''Vendor Invoice Date'' ,REFDATE ''SAP Posting Date'' " & _
        " ,''JE'' ''Vendor / Payee Name'',MEMO ''Journal Remark'' ,  [OcrCode3] ''Distribution Code'', '''' ''Purchasing Department'', Creator [Document Creator], plan_id [Expenses],  Line [Total Bill],' + @cols + '  " & _
        " from (" & _
               " SELECT i.SeriesName, c.Number   , b.REF2  ,b.TAXDATE  ,c.REFDATE  ,c.MEMO  ,  b.[OcrCode3]  , a.AcctCode + '' - '' + a.AcctName [plan_id] , e.[PrcCode] [U_AB_NONPROJECT] , j.[U_NAME] +'', '' + j.[USER_CODE][Creator] ," & _
                "   sum( round((ISNULL(b.Debit,0) - ISNULL(b.Credit,0)) /(d.[OcrTotal] /e.[PrcAmount] ),3)) AS total , sum(ISNULL(b.Debit,0) - ISNULL(b.Credit,0)) [Line] " & _
                " FROM JDT1 b JOIN OACT a ON a.AcctCode = b.Account " & _
    " JOIN OJDT c on b.[TransId] = c.[TransId] " & _
    " left join OOCR d on d.OcrCode = b.OcrCode3 " & _
    " join OCR1 e on d.[OcrCode] = e.[OcrCode] " & _
    "  LEFT JOIN NNM1 i on i.Series=c.Series " & _
    "  LEFT JOIN OUSR j on  c.[usersign] = j.[USERID]" & _
    " where  b.OcrCode3 in (" & sQuotes & ") and a.GroupMask = 5  and isnull(c.Indicator,'''') <> ''CA'' and " & _
    "	b.Account >= ''" & sSplitG(0) & "'' and b.Account <= ''" & sSplitG(1) & "'' and c.transType IN (''30'',''1470000071'') and month(c.refdate) >= " & sSplitM(0) & " and month(c.refdate) <= " & sSplitM(1) & " and year(c.refdate) = " & Now.Year & " " & _
    " and  e.[ValidFrom] <= ''" & sNow & "'' and (e.[ValidTo] >= ''" & sNow & "'' or  isnull(e.[ValidTo],'''') = '''')" & _
     " and e.PrcCode not in (SELECT T0.[U_PrcCode] FROM [dbo].[@AB_COSTALLOCATION]  T0 WHERE T0.[U_Transid] = c.TransId and T0.U_OcrCode3 = b.OcrCode3) " & _
    " group by c.Number   , b.REF2  ,b.TAXDATE  ,c.REFDATE  ,c.MEMO  ,  b.[OcrCode3]  , a.AcctCode , a.AcctName , e.[PrcCode],i.SeriesName,j.[U_NAME] , j.[USER_CODE] " & _
             "   ) x pivot ( sum(total) " & _
             "   for U_AB_NONPROJECT in (' + @cols + ') " & _
           " ) p " & _
           " UNION ALL SELECT  SeriesName, ''PS '' + cast(DOCNUM as varchar(30))''SAP Reference Number'' , '''' ''Payment Doc Num'', NULL ''Payment Doc Date'' ,'''' ''Vendor Invoice Number'',TAXDATE ''Vendor Invoice Date'',DOCDATE ''SAP Posting Date'',  " & _
    " Address ''Vendor / Payee Name'',JRNLMEMO ''Journal Remark'',  [OcrCode3] ''Distribution Code'', '''' ''Purchasing Department'', Creator [Document Creator], " & _
    "plan_id [Expenses], Line [Total Bill],' + @cols + '  " & _
                                          " from (" & _
               " SELECT i.SeriesName,f.DOCNUM   ,f.TAXDATE  ,f.DOCDATE   ,f.JRNLMEMO  ,  b.[OcrCode3]   ," & _
               "a.AcctCode + '' - '' + a.AcctName [plan_id] , e.[PrcCode] [U_AB_NONPROJECT],  f.[Address] , g.LineId,  j.[U_NAME] +'', '' + j.[USER_CODE][Creator] ," & _
                "  sum(isnull(g.SumApplied,0)) [Line] , sum( round(isnull(g.SumApplied,0) /(d.[OcrTotal] /e.[PrcAmount] ),3)) AS total " & _
                " FROM JDT1 b JOIN OACT a ON a.AcctCode = b.Account " & _
    " JOIN OJDT c on b.[TransId] = c.[TransId] " & _
    " left join OOCR d on d.OcrCode = b.OcrCode3 " & _
    " join OCR1 e on d.[OcrCode] = e.[OcrCode] " & _
    " JOIN OVPM f on c.[TransId] = f.[TransId]  JOIN VPM4 g on g.[DocNum] = f.[DocNum] and g.AcctCode = b.Account and  g.[OcrCode3] =  b.[OcrCode3] " & _
     "  LEFT JOIN NNM1 i on i.Series=f.Series " & _
      "   LEFT JOIN OUSR j on  f.[usersign] = j.[USERID] " & _
     " where  b.OcrCode3 in (" & sQuotes & ") and a.GroupMask = 5  and isnull(c.Indicator,'''') <> ''CA'' and " & _
    "	b.Account >= ''" & sSplitG(0) & "'' and b.Account <= ''" & sSplitG(1) & "'' and c.transType IN (''46'') and month(c.refdate) >= " & sSplitM(0) & " and month(c.refdate) <= " & sSplitM(1) & " and year(c.refdate) = " & Now.Year & " " & _
    " and  e.[ValidFrom] <= ''" & sNow & "'' and (e.[ValidTo] >= ''" & sNow & "'' or  isnull(e.[ValidTo],'''') = '''')" & _
     " and e.PrcCode not in (SELECT T0.[U_PrcCode] FROM [dbo].[@AB_COSTALLOCATION]  T0 WHERE T0.[U_Transid] = c.TransId and T0.U_OcrCode3 = b.OcrCode3) " & _
     "  and f.DocType = ''A'' " & _
             "  group by a.AcctCode,a.AcctName,e.[PrcCode] , f.DOCNUM ,f.TAXDATE  ,f.DOCDATE  ,f.JRNLMEMO  ,  b.[OcrCode3]  , g.LineId , f.[Address],i.SeriesName ,j.[U_NAME] , j.[USER_CODE] ) x pivot ( sum(total) " & _
             "   for U_AB_NONPROJECT in (' + @cols + ') " & _
           " ) p " & _
           " UNION ALL SELECT  SeriesName, ''PS '' + cast(DOCNUM as varchar(30))''SAP Reference Number'' , '''' ''Payment Doc Num'', NULL ''Payment Doc Date'' ,'''' ''Vendor Invoice Number'',TAXDATE ''Vendor Invoice Date'',DOCDATE ''SAP Posting Date'',  " & _
    " Address ''Vendor / Payee Name'',JRNLMEMO ''Journal Remark'',  [OcrCode3] ''Distribution Code'', '''' ''Purchasing Department'', Creator [Document Creator], " & _
    "plan_id [Expenses], Line [Total Bill],' + @cols + '  " & _
                                          " from (" & _
               " SELECT i.SeriesName,f.DOCNUM   ,f.CancelDate [TAXDATE]  ,f.CancelDate [DOCDATE]   ,f.JRNLMEMO  ,  b.[OcrCode3]   ," & _
               "a.AcctCode + '' - '' + a.AcctName [plan_id] , e.[PrcCode] [U_AB_NONPROJECT],  f.[Address] , g.LineId ,  j.[U_NAME] +'', '' + j.[USER_CODE][Creator] ," & _
                "  -sum(isnull(g.SumApplied,0)) [Line] , -sum( round(isnull(g.SumApplied,0) /(d.[OcrTotal] /e.[PrcAmount] ),3)) AS total " & _
                " FROM JDT1 b JOIN OACT a ON a.AcctCode = b.Account " & _
    " JOIN OJDT c on b.[TransId] = c.[TransId] " & _
    " left join OOCR d on d.OcrCode = b.OcrCode3 " & _
    " join OCR1 e on d.[OcrCode] = e.[OcrCode] " & _
    " JOIN OVPM f on c.[TransId] = f.[TransId]  JOIN VPM4 g on g.[DocNum] = f.[DocNum] and g.AcctCode = b.Account  and  g.[OcrCode3] =  b.[OcrCode3] " & _
     "  LEFT JOIN NNM1 i on i.Series=f.Series " & _
      "   LEFT JOIN OUSR j on  f.[usersign] = j.[USERID] " & _
     " where  b.OcrCode3 in (" & sQuotes & ") and a.GroupMask = 5  and isnull(c.Indicator,'''') <> ''CA'' and " & _
    "	b.Account >= ''" & sSplitG(0) & "'' and b.Account <= ''" & sSplitG(1) & "'' and c.transType IN (''46'') and month(f.CancelDate) >= " & sSplitM(0) & " and month(f.CancelDate) <= " & sSplitM(1) & " and year(f.CancelDate) = " & Now.Year & " " & _
    " and  e.[ValidFrom] <= ''" & sNow & "'' and (e.[ValidTo] >= ''" & sNow & "'' or  isnull(e.[ValidTo],'''') = '''')" & _
     " and e.PrcCode not in (SELECT T0.[U_PrcCode] FROM [dbo].[@AB_COSTALLOCATION]  T0 WHERE T0.[U_Transid] = c.TransId and T0.U_OcrCode3 = b.OcrCode3) " & _
     "  and f.DocType = ''A'' and f.Canceled = ''Y''" & _
             "  group by a.AcctCode,a.AcctName,e.[PrcCode] , f.DOCNUM ,f.CancelDate,f.JRNLMEMO  ,  b.[OcrCode3]  , g.LineId , f.[Address],i.SeriesName ,j.[U_NAME] , j.[USER_CODE]  ) x pivot ( sum(total) " & _
             "   for U_AB_NONPROJECT in (' + @cols + ') " & _
           " ) p ' as nvarchar(max)) " & _
  " execute (@query + @query1)"


                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Tab 2 " & sSQL, sFuncName)
                                        Try
                                            oForm.DataSources.DataTables.Add("MyDataTable1")
                                        Catch ex As Exception
                                        End Try
                                        oGridT2.DataTable = oForm.DataSources.DataTables.Item("MyDataTable1")
                                        oForm.DataSources.DataTables.Item(1).ExecuteQuery(sSQL)
                                        oGridT2.CollapseLevel = 2
                                        oGridT2.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
                                        oGridT2.AutoResizeColumns()
                                        '' oGridT2.CommonSetting.FixedColumnsCount = 3

                                        sSQL = " select * from (SELECT a.AcctCode [Account], a.AcctName [ShortName], e.U_AB_NONPROJECT [U_AB_NONPROJECT],  isnull(e.Project,'') [Project] ,e1.PrcCode [OcrCode3] , f.PrcName [U_AB_OUName]," & _
    " round((ISNULL(b.Debit,0) - ISNULL(b.Credit,0)) /(d1.[OcrTotal] /e1.[PrcAmount] ),3) AS Credit, isnull(b.LineMemo,'') [U_AB_REMARKS] , '' U_AB_PARTNER, b.TransId, d.DocEntry [BaseRef] , 'AP' 'LineMemo' " & _
                                        " , isnull(e.OcrCode2,'') [BU], isnull(e.OcrCode,'') [LOS] , isnull( b.VatGroup  ,'') [VatGroup] , b.OcrCode3 [ContraAct] " & _
    "FROM JDT1 b JOIN OACT a ON a.AcctCode = b.Account  " & _
    " JOIN OJDT c on b.[TransId] = c.[TransId]  JOIN OPCH d on c.[TransId] = d.[TransId]  " & _
    " JOIN PCH1 e on e.[DocEntry] = d.[DocEntry] and e.AcctCode = b.Account and e.[OcrCode3] = b.[OcrCode3] " & _
     " left join OOCR d1 on d1.OcrCode = b.OcrCode3 " & _
    " join OCR1 e1 on d1.[OcrCode] = e1.[OcrCode] left outer join OPRC f on f.PrcCode = e1.PrcCode " & _
    " where  b.OcrCode3 in (" & sIN & ") and a.GroupMask = 5  and isnull(d.Indicator,'') <> 'CA' and " & _
   "	b.Account >= '" & sSplitG(0) & "' and b.Account <= '" & sSplitG(1) & "' and c.transType IN ('18') and month(c.refdate) >= " & sSplitM(0) & " and month(c.refdate) <= " & sSplitM(1) & " and year(c.refdate) = " & Now.Year & " " & _
    " and  e1.[ValidFrom] <= '" & sNow & "' and (e1.[ValidTo] >= '" & sNow & "' or  isnull(e1.[ValidTo],'') = '')" & _
    " and e1.PrcCode not in (SELECT T0.[U_PrcCode] FROM [dbo].[@AB_COSTALLOCATION]  T0 WHERE T0.[U_Transid] = c.TransId and T0.U_OcrCode3 = b.OcrCode3) " & _
    " group by " & _
    "  a.AcctCode  , a.AcctName  , e.U_AB_NONPROJECT  ,  isnull(e.Project,'') ,e1.PrcCode  , f.PrcName  ,  b.Debit, b.Credit,d1.[OcrTotal],e1.[PrcAmount], isnull(b.LineMemo,'') ,  b.TransId, d.DocEntry , isnull(e.OcrCode2,'')  , isnull(e.OcrCode ,'') ,isnull( b.VatGroup  ,'') , b.OcrCode3 " & _
                                      "  UNION ALL " & _
    " SELECT   a.AcctCode [Account], a.AcctName [ShortName], e.U_AB_NONPROJECT [U_AB_NONPROJECT],  isnull(e.Project,'') [Project], e1.PrcCode [OcrCode3] , f.PrcName [U_AB_OUName] , " & _
    " round((ISNULL(b.Debit,0) - ISNULL(b.Credit,0)) /(d1.[OcrTotal] /e1.[PrcAmount] ),3) AS Credit,isnull( b.LineMemo,'') [U_AB_REMARKS] , '' U_AB_PARTNER ,  b.TransId, d.DocEntry [BaseRef] , 'CN' 'LineMemo'  " & _
    " , isnull(e.OcrCode2,'') [BU], isnull(e.OcrCode,'') [LOS] , isnull( b.VatGroup  ,'') [VatGroup], b.OcrCode3 [ContraAct] " & _
    "  FROM JDT1 b JOIN OACT a  ON a.AcctCode = b.Account  JOIN OJDT c on b.[TransId] = c.[TransId]  " & _
    " JOIN ORPC d on c.[TransId] = d.[TransId]  JOIN RPC1 e on e.[DocEntry] = d.[DocEntry] and e.AcctCode = b.Account  and e.[OcrCode3] = b.[OcrCode3]  " & _
     " left join OOCR d1 on d1.OcrCode = b.OcrCode3 " & _
    " join OCR1 e1 on d1.[OcrCode] = e1.[OcrCode]  left outer join OPRC f on f.PrcCode = e1.PrcCode" & _
    " where  b.OcrCode3 in (" & sIN & ") and a.GroupMask = 5  and isnull(d.Indicator,'') <> 'CA' and " & _
   "	b.Account >= '" & sSplitG(0) & "' and b.Account <= '" & sSplitG(1) & "' and c.transType IN ('19') and month(c.refdate) >= " & sSplitM(0) & " and month(c.refdate) <= " & sSplitM(1) & " and year(c.refdate) = " & Now.Year & " " & _
    " and  e1.[ValidFrom] <= '" & sNow & "' and (e1.[ValidTo] >= '" & sNow & "' or  isnull(e1.[ValidTo],'') = '')" & _
 " and e1.PrcCode not in (SELECT T0.[U_PrcCode] FROM [dbo].[@AB_COSTALLOCATION]  T0 WHERE T0.[U_Transid] = c.TransId and T0.U_OcrCode3 = b.OcrCode3) " & _
     " group by " & _
    "  a.AcctCode  , a.AcctName  , e.U_AB_NONPROJECT  , isnull( e.Project,'') ,e1.PrcCode  , f.PrcName  ,  b.Debit, b.Credit,d1.[OcrTotal],e1.[PrcAmount],isnull( b.LineMemo,'') ,  b.TransId, d.DocEntry  , isnull(e.OcrCode2,'')  , isnull(e.OcrCode ,'') ,isnull( b.VatGroup  ,''), b.OcrCode3 " & _
       "  UNION ALL " & _
    " SELECT   a.AcctCode [Account], a.AcctName [ShortName], b.U_AB_NONPROJECT [U_AB_NONPROJECT],  isnull(b.Project,'') [Project] ,e1.PrcCode [OcrCode3] , f.PrcName [U_AB_OUName] , " & _
    " round((ISNULL(b.Debit,0) - ISNULL(b.Credit,0)) /(d1.[OcrTotal] /e1.[PrcAmount] ),3) AS Credit, isnull(b.LineMemo,'') [U_AB_REMARKS] ,  isnull(g.Name,'')  U_AB_PARTNER ,  b.TransId, 0 [BaseRef] , case when c.transType = '30' then case when left(a.AcctCode,4) = '7212' or left(a.AcctCode,4) = '7213' then  'DEP' else 'JE' end  else 'DEP' end 'LineMemo', isnull(b.OcrCode2,'') [BU] , isnull(b.ProfitCode,'') [LOS] , isnull( b.VatGroup  ,'') [VatGroup] ,b.OcrCode3 [ContraAct] " & _
    "  FROM JDT1 b JOIN OACT a  ON a.AcctCode = b.Account  JOIN OJDT c on b.[TransId] = c.[TransId]  " & _
    " left join OOCR d1 on d1.OcrCode = b.OcrCode3 " & _
    " join OCR1 e1 on d1.[OcrCode] = e1.[OcrCode]  left outer join OPRC f on f.PrcCode = e1.PrcCode" & _
    " left outer join [@AB_PARTNER] g on g.Code = b.U_AB_PARTNER" & _
    " where  b.OcrCode3 in (" & sIN & ") and a.GroupMask = 5  and isnull(c.Indicator,'') <> 'CA' and " & _
   "	b.Account >= '" & sSplitG(0) & "' and b.Account <= '" & sSplitG(1) & "' and c.transType IN ('30', '1470000071') and month(c.refdate) >= " & sSplitM(0) & " and month(c.refdate) <= " & sSplitM(1) & " and year(c.refdate) = " & Now.Year & " " & _
    " and  e1.[ValidFrom] <= '" & sNow & "' and (e1.[ValidTo] >= '" & sNow & "' or  isnull(e1.[ValidTo],'') = '')" & _
     " and e1.PrcCode not in (SELECT T0.[U_PrcCode] FROM [dbo].[@AB_COSTALLOCATION]  T0 WHERE T0.[U_Transid] = c.TransId and T0.U_OcrCode3 = b.OcrCode3) " & _
     " group by " & _
    "  a.AcctCode  , a.AcctName  , b.U_AB_NONPROJECT  , isnull( b.Project,'') ,e1.PrcCode  , f.PrcName  ,  b.Debit, b.Credit,d1.[OcrTotal],e1.[PrcAmount],isnull( b.LineMemo,'') ,  b.TransId, isnull( b.OcrCode2,'')   , isnull(b.ProfitCode,''), isnull(g.Name,''),isnull( b.VatGroup  ,'') , b.OcrCode3 ,c.transType" & _
    " UNION ALL SELECT a.AcctCode [Account], a.AcctName [ShortName], e.U_AB_NONPROJECT [U_AB_NONPROJECT],  isnull(e.Project,'') [Project] ,e1.PrcCode [OcrCode3] , f.PrcName [U_AB_OUName]," & _
    " round((ISNULL(b.Debit,0) - ISNULL(b.Credit,0)) /(d1.[OcrTotal] /e1.[PrcAmount] ),3) AS Credit, isnull(b.LineMemo,'') [U_AB_REMARKS] , '' U_AB_PARTNER, b.TransId, d.DocEntry [BaseRef] , 'AP' 'LineMemo' " & _
                                        " , isnull(e.OcrCode2,'') [BU], isnull(e.OcrCode,'') [LOS] , isnull( b.VatGroup  ,'') [VatGroup] , b.OcrCode3 [ContraAct] " & _
    "FROM JDT1 b JOIN OACT a ON a.AcctCode = b.Account  " & _
    " JOIN OJDT c on b.[TransId] = c.[TransId]  JOIN OVPM d on c.[TransId] = d.[TransId]  " & _
    " JOIN VPM4 e on e.[DocNum] = d.[DocNum] and e.AcctCode = b.Account  and e.[OcrCode3] = b.[OcrCode3] " & _
     " left join OOCR d1 on d1.OcrCode = b.OcrCode3 " & _
    " join OCR1 e1 on d1.[OcrCode] = e1.[OcrCode] left outer join OPRC f on f.PrcCode = e1.PrcCode " & _
    " where  b.OcrCode3 in (" & sIN & ") and a.GroupMask = 5  and d.DocType = 'A' and " & _
   "	b.Account >= '" & sSplitG(0) & "' and b.Account <= '" & sSplitG(1) & "' and c.transType IN ('46') and month(c.refdate) >= " & sSplitM(0) & " and month(c.refdate) <= " & sSplitM(1) & " and year(c.refdate) = " & Now.Year & " " & _
    " and  e1.[ValidFrom] <= '" & sNow & "' and (e1.[ValidTo] >= '" & sNow & "' or  isnull(e1.[ValidTo],'') = '')" & _
    " and e1.PrcCode not in (SELECT T0.[U_PrcCode] FROM [dbo].[@AB_COSTALLOCATION]  T0 WHERE T0.[U_Transid] = c.TransId and T0.U_OcrCode3 = b.OcrCode3) " & _
    " group by " & _
    "  a.AcctCode  , a.AcctName  , e.U_AB_NONPROJECT  ,  isnull(e.Project,'') ,e1.PrcCode  , f.PrcName  ,  b.Debit, b.Credit,d1.[OcrTotal],e1.[PrcAmount], isnull(b.LineMemo,'') ,  b.TransId, d.DocEntry , isnull(e.OcrCode2,'')  , isnull(e.OcrCode ,'') ,isnull( b.VatGroup  ,''), b.OcrCode3 " & _
   " UNION ALL SELECT a.AcctCode [Account], a.AcctName [ShortName], e.U_AB_NONPROJECT [U_AB_NONPROJECT],  isnull(e.Project,'') [Project] ,e1.PrcCode [OcrCode3] , f.PrcName [U_AB_OUName]," & _
    " -round((ISNULL(b.Debit,0) - ISNULL(b.Credit,0)) /(d1.[OcrTotal] /e1.[PrcAmount] ),3) AS Credit, isnull(b.LineMemo,'') [U_AB_REMARKS] , '' U_AB_PARTNER, b.TransId, d.DocEntry [BaseRef] , 'AP' 'LineMemo' " & _
                                        " , isnull(e.OcrCode2,'') [BU], isnull(e.OcrCode,'') [LOS] , isnull( b.VatGroup  ,'') [VatGroup] , b.OcrCode3 [ContraAct] " & _
    "FROM JDT1 b JOIN OACT a ON a.AcctCode = b.Account  " & _
    " JOIN OJDT c on b.[TransId] = c.[TransId]  JOIN OVPM d on c.[TransId] = d.[TransId]  " & _
    " JOIN VPM4 e on e.[DocNum] = d.[DocNum] and e.AcctCode = b.Account  and e.[OcrCode3] = b.[OcrCode3]  " & _
     " left join OOCR d1 on d1.OcrCode = b.OcrCode3 " & _
    " join OCR1 e1 on d1.[OcrCode] = e1.[OcrCode] left outer join OPRC f on f.PrcCode = e1.PrcCode " & _
    " where  b.OcrCode3 in (" & sIN & ") and a.GroupMask = 5  and d.DocType = 'A' and d.Canceled = 'Y' and " & _
   "	b.Account >= '" & sSplitG(0) & "' and b.Account <= '" & sSplitG(1) & "' and c.transType IN ('46') and month(d.CancelDate) >= " & sSplitM(0) & " and month(d.CancelDate) <= " & sSplitM(1) & " and year(d.CancelDate) = " & Now.Year & " " & _
    " and  e1.[ValidFrom] <= '" & sNow & "' and (e1.[ValidTo] >= '" & sNow & "' or  isnull(e1.[ValidTo],'') = '')" & _
    " and e1.PrcCode not in (SELECT T0.[U_PrcCode] FROM [dbo].[@AB_COSTALLOCATION]  T0 WHERE T0.[U_Transid] = c.TransId and T0.U_OcrCode3 = b.OcrCode3) " & _
    " group by " & _
    "  a.AcctCode  , a.AcctName  , e.U_AB_NONPROJECT  ,  isnull(e.Project,'') ,e1.PrcCode  , f.PrcName  ,  b.Debit, b.Credit,d1.[OcrTotal],e1.[PrcAmount], isnull(b.LineMemo,'') ,  b.TransId, d.DocEntry , isnull(e.OcrCode2,'')  , isnull(e.OcrCode ,'') ,isnull( b.VatGroup  ,''), b.OcrCode3 " & _
                                        ")x order by x.OcrCode3 , cast(x.Account as integer)"

                                        Try
                                            oForm.DataSources.DataTables.Add("JDT1")
                                        Catch ex As Exception
                                        End Try
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Tab 3 " & sSQL, sFuncName)
                                        Dim oSApDT As SAPbouiCOM.DataTable

                                        oSApDT = oForm.DataSources.DataTables.Item("JDT1")
                                        '' oForm.DataSources.DataTables.Item("JDT1").ExecuteQuery(sSQL)
                                        oSApDT.ExecuteQuery(sSQL)
                                        oMatrix.CommonSetting.SeparateLine(oMatrix.RowCount, 245, SAPbouiCOM.BoSeparateLineType.slt_Bottom, False)
                                        oMatrix.CommonSetting.SeparateLine(oMatrix.RowCount - 1, 245, SAPbouiCOM.BoSeparateLineType.slt_Bottom, False)
                                        oMatrix.CommonSetting.SeparateLine(oMatrix.RowCount - 2, 245, SAPbouiCOM.BoSeparateLineType.slt_Bottom, False)
                                        '' oMatrix.CommonSetting.SetRowFontSize(oMatrix.RowCount, 9)
                                        oMatrix.Clear()

                                        oForm.Items.Item("Item_24").Specific.columns.item("Col_0").databind.bind("JDT1", "Account")
                                        oForm.Items.Item("Item_24").Specific.columns.item("Col_1").databind.bind("JDT1", "ShortName")
                                        oForm.Items.Item("Item_24").Specific.columns.item("Col_2").databind.bind("JDT1", "U_AB_NONPROJECT")
                                        oForm.Items.Item("Item_24").Specific.columns.item("Col_3").databind.bind("JDT1", "Project")
                                        oForm.Items.Item("Item_24").Specific.columns.item("Col_4").databind.bind("JDT1", "OcrCode3")
                                        oForm.Items.Item("Item_24").Specific.columns.item("Col_5").databind.bind("JDT1", "U_AB_OUName")
                                        ''oForm.Items.Item("Item_24").Specific.columns.item("Col_6").databind.bind("JDT1", "U_AB_PARTNER")
                                        oForm.Items.Item("Item_24").Specific.columns.item("Col_7").databind.bind("JDT1", "Credit")
                                        oForm.Items.Item("Item_24").Specific.columns.item("Col_8").databind.bind("JDT1", "U_AB_REMARKS")
                                        oForm.Items.Item("Item_24").Specific.columns.item("Col_9").databind.bind("JDT1", "TransId")
                                        oForm.Items.Item("Item_24").Specific.columns.item("Col_10").databind.bind("JDT1", "BaseRef")
                                        oForm.Items.Item("Item_24").Specific.columns.item("Col_11").databind.bind("JDT1", "LineMemo")
                                        oForm.Items.Item("Item_24").Specific.columns.item("V_1").databind.bind("JDT1", "BU")
                                        oForm.Items.Item("Item_24").Specific.columns.item("V_0").databind.bind("JDT1", "LOS")
                                        oForm.Items.Item("Item_24").Specific.columns.item("V_3").databind.bind("JDT1", "U_AB_PARTNER")
                                        oForm.Items.Item("Item_24").Specific.columns.item("V_2").databind.bind("JDT1", "VatGroup")
                                        oForm.Items.Item("Item_24").Specific.columns.item("V_4").databind.bind("JDT1", "ContraAct")
                                        oForm.Items.Item("Item_24").Specific.LoadFromDataSource()
                                        oForm.Items.Item("Item_24").Specific.AutoResizeColumns()
                                        oMatrix.Columns.Item("Col_6").Editable = True

                                        Dim dsum As Double = 0
                                        For imjs As Integer = 0 To oSApDT.Rows.Count - 1
                                            dsum += oSApDT.GetValue("Credit", imjs)
                                        Next
                                        oForm.DataSources.DataTables.Item(2).Clear()

                                        oMatrix.AddRow(2)
                                        oMatrix.CommonSetting.SeparateLine(oMatrix.RowCount, 245, SAPbouiCOM.BoSeparateLineType.slt_Bottom, True)
                                        oMatrix.AddRow(1)
                                        oMatrix.Columns.Item("Col_7").Cells.Item(oMatrix.RowCount).Specific.String = dsum
                                        ''oMatrix.CommonSetting.SetRowFontSize(oMatrix.RowCount, 11)
                                        oMatrix.Columns.Item("Col_5").Cells.Item(oMatrix.RowCount).Specific.String = "Total"
                                        oMatrix.CommonSetting.SeparateLine(oMatrix.RowCount, 245, SAPbouiCOM.BoSeparateLineType.slt_Bottom, True)


                                        oForm.Freeze(False)
                                        p_oSBOApplication.SetStatusBarMessage("Cost Allocation information successfully loaded .......!", SAPbouiCOM.BoMessageTime.bmt_Short, False)

                                    Catch ex As Exception
                                        oForm.Freeze(False)
                                        p_oSBOApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                        BubbleEvent = False
                                        sErrDesc = ex.Message
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                                        WriteToLogFile(Err.Description, sFuncName)
                                        ShowErr(sErrDesc)
                                    End Try
                                End If


                                If pVal.ItemUID = "Item_21" Then
                                    Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                                    Dim sMonth As String = String.Empty
                                    Dim sSplitM() As String
                                    Dim sSplitD() As String
                                    Dim sSplitG() As String
                                    Dim sDistribution As String = String.Empty
                                    Dim sGLAccount As String = String.Empty
                                    Dim sSQL As String = String.Empty
                                    Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("Item_24").Specific
                                    Dim sNow As String = String.Empty
                                    Dim sIN As String = String.Empty
                                    Dim sQuotes As String = String.Empty
                                    Dim oRset As SAPbobsCOM.Recordset = Nothing
                                    Dim sCostF As String = String.Empty
                                    Dim sCostT As String = String.Empty
                                    Dim oGridT2 As SAPbouiCOM.Grid = Nothing

                                    Try

                                        oRset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        sFuncName = "Journal Entry Creation Show button click()"
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling HeaderValidation()", sFuncName)
                                        If CostAllocation_Validation(oForm, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                        If String.IsNullOrEmpty(oForm.Items.Item("Item_2").Specific.String) Or String.IsNullOrEmpty(oForm.Items.Item("Item_15").Specific.String) Then
                                            p_oSBOApplication.SetStatusBarMessage("Operating Unit should not be Empty ...... !", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                        oForm.Freeze(True)
                                        sMonth = oForm.Items.Item("Item_13").Specific.selected.description & "," & oForm.Items.Item("Item_14").Specific.selected.description
                                        sDistribution = oForm.Items.Item("Item_5").Specific.value & "," & oForm.Items.Item("Item_8").Specific.value
                                        sGLAccount = oForm.Items.Item("Item_10").Specific.value & "," & oForm.Items.Item("Item_12").Specific.value
                                        sSplitM = sMonth.Split(",")
                                        sSplitD = sDistribution.Split(",")
                                        sSplitG = sGLAccount.Split(",")
                                        sCostF = oForm.Items.Item("Item_2").Specific.String
                                        sCostT = oForm.Items.Item("Item_15").Specific.String

                                        oForm.Items.Item("Item_22").Specific.String() = sCostF
                                        oForm.Items.Item("Item_26").Specific.String = sCostT

                                        sSQL = "DECLARE  " & _
                                         "  @string varchar(100), @string1 varchar(100), @string2 varchar(100), @string3 varchar(max), @string4 varchar(max) " & _
                                         "  SET @string = '" & sSplitD(0) & "' SET @string1 = '" & sSplitD(1) & "' SET @string2 = '" & sSplitD(0) & "' " & _
                                         " WHILE PATINDEX('%[^a-z]%',@string2) > 0 SET @string2 = STUFF(@string2,PATINDEX('%[^a-z]%',@string2),1,'') " & _
                                         " WHILE PATINDEX('%[^0-9]%',@string) <> 0     SET @string = STUFF(@string,PATINDEX('%[^0-9]%',@string),1,'') " & _
                                         " WHILE PATINDEX('%[^0-9]%',@string1) <> 0     SET @string1 = STUFF(@string1,PATINDEX('%[^0-9]%',@string1),1,'') " & _
                                         " WHILE cast(@string as numeric ) <= cast(@string1 as numeric ) " & _
                                         " begin " & _
                                         " SET @string3 =  isnull(@string3,'')  + '[' + @string2 + @string + '],' " & _
                                         " set @string = cast(@string as numeric ) + 1 " & _
                                         " end " & _
                                         " set @string3 = replace( replace(@string3,'[',''''),']','''') " & _
                                         " set @string3 = left(@string3, len(@string3) -1) " & _
                                           " set @string4 = replace(@string3,'''','''''') " & _
                                         " select @string3 [Ouput] , @string4 [quotes] "
                                        oRset.DoQuery(sSQL)
                                        sIN = oRset.Fields.Item("Ouput").Value
                                        sQuotes = oRset.Fields.Item("quotes").Value
                                        sNow = CStr(Now.Year) & sSplitM(0).PadLeft(2, "0"c) & "01"


                                        sSQL = "SELECT a.AcctCode [Account], a.AcctName [ShortName], e.U_AB_NONPROJECT [U_AB_NONPROJECT],  isnull(e.Project,'') [Project] ,e1.PrcCode [OcrCode3] , f.PrcName [U_AB_OUName]," & _
                                        " round((ISNULL(b.Debit,0) - ISNULL(b.Credit,0)) /(d1.[OcrTotal] /e1.[PrcAmount] ),3) AS Credit, isnull(b.LineMemo,'') [U_AB_REMARKS] , '' U_AB_PARTNER, b.TransId, d.DocEntry [BaseRef] , 'AP' 'LineMemo' " & _
                                        " , isnull(e.OcrCode2,'') [BU], isnull(e.OcrCode,'') [LOS] , isnull( b.VatGroup  ,'') [VatGroup] " & _
                                          "FROM JDT1 b JOIN OACT a ON a.AcctCode = b.Account  " & _
                                          " JOIN OJDT c on b.[TransId] = c.[TransId]  JOIN OPCH d on c.[TransId] = d.[TransId]  " & _
                                          " JOIN PCH1 e on e.[DocEntry] = d.[DocEntry] " & _
                                           " left join OOCR d1 on d1.OcrCode = b.OcrCode3 " & _
                                          " join OCR1 e1 on d1.[OcrCode] = e1.[OcrCode] left outer join OPRC f on f.PrcCode = e1.PrcCode " & _
                                          " where  b.OcrCode3 in (" & sIN & ") and a.GroupMask = 5  and isnull(d.Indicator,'') <> 'CA' and " & _
                                         "	b.Account >= '" & sSplitG(0) & "' and b.Account <= '" & sSplitG(1) & "' and c.transType IN ('18') and month(c.refdate) >= " & sSplitM(0) & " and month(c.refdate) <= " & sSplitM(1) & " and year(c.refdate) = " & Now.Year & " " & _
                                          " and  e1.[ValidFrom] <= '" & sNow & "' and (e1.[ValidTo] >= '" & sNow & "' or  isnull(e1.[ValidTo],'') = '')" & _
                                           " and e1.PrcCode >= '" & sCostF & "' and e1.PrcCode <= '" & sCostT & "'" & _
                                            " group by " & _
                                           "  a.AcctCode  , a.AcctName  , e.U_AB_NONPROJECT  ,  isnull(e.Project,'') ,e1.PrcCode  , f.PrcName  ,  b.Debit, b.Credit,d1.[OcrTotal],e1.[PrcAmount], isnull(b.LineMemo,'') ,  b.TransId, d.DocEntry , isnull(e.OcrCode2,'')  , isnull(e.OcrCode ,'') ,isnull( b.VatGroup  ,'') " & _
                                                                               "  UNION ALL " & _
                                          " SELECT   a.AcctCode [Account], a.AcctName [ShortName], e.U_AB_NONPROJECT [U_AB_NONPROJECT],  isnull(e.Project,'') [Project], e1.PrcCode [OcrCode3] , f.PrcName [U_AB_OUName] , " & _
                                          " round((ISNULL(b.Debit,0) - ISNULL(b.Credit,0)) /(d1.[OcrTotal] /e1.[PrcAmount] ),3) AS Credit,isnull( b.LineMemo,'') [U_AB_REMARKS] , '' U_AB_PARTNER ,  b.TransId, d.DocEntry [BaseRef] , 'CN' 'LineMemo'  " & _
                                          " , isnull(e.OcrCode2,'') [BU], isnull(e.OcrCode,'') [LOS] , isnull( b.VatGroup  ,'') [VatGroup] " & _
                                          "  FROM JDT1 b JOIN OACT a  ON a.AcctCode = b.Account  JOIN OJDT c on b.[TransId] = c.[TransId]  " & _
                                          " JOIN ORPC d on c.[TransId] = d.[TransId]  JOIN RPC1 e on e.[DocEntry] = d.[DocEntry]  " & _
                                           " left join OOCR d1 on d1.OcrCode = b.OcrCode3 " & _
                                          " join OCR1 e1 on d1.[OcrCode] = e1.[OcrCode]  left outer join OPRC f on f.PrcCode = e1.PrcCode" & _
                                          " where  b.OcrCode3 in (" & sIN & ") and a.GroupMask = 5  and isnull(d.Indicator,'') <> 'CA' and " & _
                                         "	b.Account >= '" & sSplitG(0) & "' and b.Account <= '" & sSplitG(1) & "' and c.transType IN ('19') and month(c.refdate) >= " & sSplitM(0) & " and month(c.refdate) <= " & sSplitM(1) & " and year(c.refdate) = " & Now.Year & " " & _
                                          " and  e1.[ValidFrom] <= '" & sNow & "' and (e1.[ValidTo] >= '" & sNow & "' or  isnull(e1.[ValidTo],'') = '')" & _
                                           " and e1.PrcCode >= '" & sCostF & "' and e1.PrcCode <= '" & sCostT & "'" & _
                                            " group by " & _
                                          "  a.AcctCode  , a.AcctName  , e.U_AB_NONPROJECT  , isnull( e.Project,'') ,e1.PrcCode  , f.PrcName  ,  b.Debit, b.Credit,d1.[OcrTotal],e1.[PrcAmount],isnull( b.LineMemo,'') ,  b.TransId, d.DocEntry  , isnull(e.OcrCode2,'')  , isnull(e.OcrCode ,'') ,isnull( b.VatGroup  ,'') " & _
                                                                                  "  UNION ALL " & _
                                         " SELECT   a.AcctCode [Account], a.AcctName [ShortName], b.U_AB_NONPROJECT [U_AB_NONPROJECT],  isnull(b.Project,'') [Project] ,e1.PrcCode [OcrCode3] , f.PrcName [U_AB_OUName] , " & _
                                         " round((ISNULL(b.Debit,0) - ISNULL(b.Credit,0)) /(d1.[OcrTotal] /e1.[PrcAmount] ),3) AS Credit, isnull(b.LineMemo,'') [U_AB_REMARKS] ,  isnull(g.Name,'')  U_AB_PARTNER ,  b.TransId, 0 [BaseRef] , case when c.transType = '30' then case when left(a.AcctCode,4) = '7212' or left(a.AcctCode,4) = '7213' then  'DEP' else 'JE' end  else 'DEP' end 'LineMemo', isnull(b.OcrCode2,'') [BU] , isnull(b.ProfitCode,'') [LOS] , isnull( b.VatGroup  ,'') [VatGroup]  " & _
                                          "  FROM JDT1 b JOIN OACT a  ON a.AcctCode = b.Account  JOIN OJDT c on b.[TransId] = c.[TransId]  " & _
                                          " left join OOCR d1 on d1.OcrCode = b.OcrCode3 " & _
                                          " join OCR1 e1 on d1.[OcrCode] = e1.[OcrCode]  left outer join OPRC f on f.PrcCode = e1.PrcCode" & _
                                          " left outer join [@AB_PARTNER] g on g.Code = b.U_AB_PARTNER" & _
                                          " where  b.OcrCode3 in (" & sIN & ") and a.GroupMask = 5  and isnull(c.Indicator,'') <> 'CA' and " & _
                                         "	b.Account >= '" & sSplitG(0) & "' and b.Account <= '" & sSplitG(1) & "' and c.transType IN ('30','1470000071') and month(c.refdate) >= " & sSplitM(0) & " and month(c.refdate) <= " & sSplitM(1) & " and year(c.refdate) = " & Now.Year & " " & _
                                          " and  e1.[ValidFrom] <= '" & sNow & "' and (e1.[ValidTo] >= '" & sNow & "' or  isnull(e1.[ValidTo],'') = '')" & _
                                           " and e1.PrcCode >= '" & sCostF & "' and e1.PrcCode <= '" & sCostT & "' " & _
                                            " group by " & _
                                           "  a.AcctCode  , a.AcctName  , b.U_AB_NONPROJECT  , isnull( b.Project,'') ,e1.PrcCode  , f.PrcName  ,  b.Debit, b.Credit,d1.[OcrTotal],e1.[PrcAmount],isnull( b.LineMemo,'') ,  b.TransId, isnull( b.OcrCode2,'')   , isnull(b.ProfitCode,''), isnull(g.Name,''),isnull( b.VatGroup  ,''), c.transType"


                                        Try
                                            oForm.DataSources.DataTables.Add("JDT1")
                                        Catch ex As Exception
                                        End Try
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Tab 3 - Show butto click " & sSQL, sFuncName)
                                        Dim oSApDT As SAPbouiCOM.DataTable

                                        oSApDT = oForm.DataSources.DataTables.Item("JDT1")
                                        '' oForm.DataSources.DataTables.Item("JDT1").ExecuteQuery(sSQL)
                                        oSApDT.ExecuteQuery(sSQL)
                                        oMatrix.CommonSetting.SeparateLine(oMatrix.RowCount, 245, SAPbouiCOM.BoSeparateLineType.slt_Bottom, False)
                                        oMatrix.CommonSetting.SeparateLine(oMatrix.RowCount - 1, 245, SAPbouiCOM.BoSeparateLineType.slt_Bottom, False)
                                        oMatrix.CommonSetting.SeparateLine(oMatrix.RowCount - 2, 245, SAPbouiCOM.BoSeparateLineType.slt_Bottom, False)
                                        '' oMatrix.CommonSetting.SetRowFontSize(oMatrix.RowCount, 9)
                                        oMatrix.Clear()

                                        oForm.Items.Item("Item_24").Specific.columns.item("Col_0").databind.bind("JDT1", "Account")
                                        oForm.Items.Item("Item_24").Specific.columns.item("Col_1").databind.bind("JDT1", "ShortName")
                                        oForm.Items.Item("Item_24").Specific.columns.item("Col_2").databind.bind("JDT1", "U_AB_NONPROJECT")
                                        oForm.Items.Item("Item_24").Specific.columns.item("Col_3").databind.bind("JDT1", "Project")
                                        oForm.Items.Item("Item_24").Specific.columns.item("Col_4").databind.bind("JDT1", "OcrCode3")
                                        oForm.Items.Item("Item_24").Specific.columns.item("Col_5").databind.bind("JDT1", "U_AB_OUName")
                                        ''oForm.Items.Item("Item_24").Specific.columns.item("Col_6").databind.bind("JDT1", "U_AB_PARTNER")
                                        oForm.Items.Item("Item_24").Specific.columns.item("Col_7").databind.bind("JDT1", "Credit")
                                        oForm.Items.Item("Item_24").Specific.columns.item("Col_8").databind.bind("JDT1", "U_AB_REMARKS")
                                        oForm.Items.Item("Item_24").Specific.columns.item("Col_9").databind.bind("JDT1", "TransId")
                                        oForm.Items.Item("Item_24").Specific.columns.item("Col_10").databind.bind("JDT1", "BaseRef")
                                        oForm.Items.Item("Item_24").Specific.columns.item("Col_11").databind.bind("JDT1", "LineMemo")
                                        oForm.Items.Item("Item_24").Specific.columns.item("V_1").databind.bind("JDT1", "BU")
                                        oForm.Items.Item("Item_24").Specific.columns.item("V_0").databind.bind("JDT1", "LOS")
                                        oForm.Items.Item("Item_24").Specific.columns.item("V_3").databind.bind("JDT1", "U_AB_PARTNER")
                                        oForm.Items.Item("Item_24").Specific.columns.item("V_2").databind.bind("JDT1", "VatGroup")
                                        oForm.Items.Item("Item_24").Specific.LoadFromDataSource()
                                        oForm.Items.Item("Item_24").Specific.AutoResizeColumns()
                                        oMatrix.Columns.Item("Col_6").Editable = True

                                        Dim dsum As Double = 0
                                        For imjs As Integer = 0 To oSApDT.Rows.Count - 1
                                            dsum += Math.Abs(oSApDT.GetValue("Credit", imjs))
                                        Next
                                        oForm.DataSources.DataTables.Item(2).Clear()

                                        oMatrix.AddRow(2)
                                        oMatrix.CommonSetting.SeparateLine(oMatrix.RowCount, 245, SAPbouiCOM.BoSeparateLineType.slt_Bottom, True)
                                        oMatrix.AddRow(1)
                                        oMatrix.Columns.Item("Col_7").Cells.Item(oMatrix.RowCount).Specific.String = Math.Round(dsum)
                                        oMatrix.Columns.Item("Col_5").Cells.Item(oMatrix.RowCount).Specific.String = "Total"
                                        ''  oMatrix.CommonSetting.SetRowFontSize(oMatrix.RowCount, 11)
                                        oMatrix.CommonSetting.SeparateLine(oMatrix.RowCount, 245, SAPbouiCOM.BoSeparateLineType.slt_Bottom, True)




                                        sSQL = "DECLARE @cols AS NVARCHAR(MAX),    @query  AS VARCHAR(max), @cols1 as nvarchar(max) " & _
                                              "select @cols = STUFF((SELECT distinct  ',' + QUOTENAME(cast(T1.PrcCode as nvarchar(100))) " & _
                   " from OOCR T0  INNER JOIN OCR1 T1 ON T0.[OcrCode] = T1.[OcrCode]  where T0.OcrCode in (" & sIN & ") and T1.[ValidFrom] <= '" & sNow & "' and  (T1.[ValidTo] >= '" & sNow & "' or  isnull(T1.[ValidTo],'') = '')" & _
                   " and T1.PrcCode >= '" & sCostF & "' and T1.PrcCode <= '" & sCostT & "' " & _
           " FOR XML PATH(''), TYPE " & _
           " ).value('.', 'NVARCHAR(MAX)') " & _
       "  ,1,1,'') " & _
    "select @cols1 = STUFF((SELECT '+ isnull(' + QUOTENAME(isnull(T1.PrcCode,0)) + ',0)' " & _
                   " from OOCR T0  INNER JOIN OCR1 T1 ON T0.[OcrCode] = T1.[OcrCode]  where T0.OcrCode in (" & sIN & ") and T1.[ValidFrom] <= '" & sNow & "' and  (T1.[ValidTo] >= '" & sNow & "' or  isnull(T1.[ValidTo],'') = '')" & _
                    " and T1.PrcCode >= '" & sCostF & "' and T1.PrcCode <= '" & sCostT & "' " & _
           " FOR XML PATH(''), TYPE " & _
           " ).value('.', 'NVARCHAR(MAX)') " & _
       "  ,1,1,'') " & _
    " set @query = cast('SELECT cast(DOCNUM as nvarchar(30)) ''SAP Reference Number'' , NUMATCARD ''Vendor Invoice Number'',TAXDATE ''Vendor Invoice Date'',DOCDATE ''SAP Posting Date'',  " & _
    " CARDNAME ''Vendor / Payee Name'',JRNLMEMO ''Journal Remark'',  [OcrCode3] ''Distribution Code'', SLPNAME ''Purchasing Department'', " & _
    "plan_id [Expenses], Line [Total Bill],' + @cols + '  " & _
                                          " from (" & _
               " SELECT f.DOCNUM   , f.NUMATCARD  ,f.TAXDATE  ,f.DOCDATE  ,f.CARDNAME  ,f.JRNLMEMO  ,  b.[OcrCode3]  ,h.SLPNAME  ," & _
               "a.AcctCode + '' - '' + a.AcctName [plan_id] , e.[PrcCode] [U_AB_NONPROJECT], " & _
                "  sum(ISNULL( g.LineTotal,0)) [Line] , sum( round((isnull( g.LineTotal,0)) /(d.[OcrTotal] /e.[PrcAmount] ),3)) AS total  " & _
                " FROM JDT1 b JOIN OACT a ON a.AcctCode = b.Account " & _
    " JOIN OJDT c on b.[TransId] = c.[TransId] " & _
    " left join OOCR d on d.OcrCode = b.OcrCode3 " & _
    " join OCR1 e on d.[OcrCode] = e.[OcrCode] " & _
    " JOIN OPCH f on c.[TransId] = f.[TransId]  JOIN PCH1 g on g.[DocEntry] = f.[DocEntry] and g.AcctCode = b.Account  " & _
    "  LEFT JOIN OSLP h on f.SLPCODE=h.SLPCODE " & _
    " where  b.OcrCode3 in (" & sQuotes & ") and a.GroupMask = 5  and isnull(c.Indicator,'''') <> ''CA'' and " & _
    "	b.Account >= ''" & sSplitG(0) & "'' and b.Account <= ''" & sSplitG(1) & "'' and c.transType IN (''18'') and month(c.refdate) >= " & sSplitM(0) & " and month(c.refdate) <= " & sSplitM(1) & " and year(c.refdate) = " & Now.Year & " " & _
    " and  e.[ValidFrom] <= ''" & sNow & "'' and (e.[ValidTo] >= ''" & sNow & "'' or  isnull(e.[ValidTo],'''') = '''')" & _
             "  group by a.AcctCode,a.AcctName,e.[PrcCode] , f.DOCNUM   , f.NUMATCARD  ,f.TAXDATE  ,f.DOCDATE  ,f.CARDNAME  ,f.JRNLMEMO  ,  b.[OcrCode3]  ,h.SLPNAME, g.LineNum   ) x pivot ( sum(total) " & _
             "   for U_AB_NONPROJECT in (' + @cols + ') " & _
           " ) p  " & _
               " union all SELECT cast( DOCNUM as nvarchar(30)) ''SAP Reference Number'' , NUMATCARD ''Vendor Invoice Number'',TAXDATE ''Vendor Invoice Date'',DOCDATE ''SAP Posting Date'',  " & _
    " CARDNAME ''Vendor / Payee Name'',JRNLMEMO ''Journal Remark'',  [OcrCode3] ''Distribution Code'', SLPNAME ''Purchasing Department'', " & _
    "plan_id [Expenses], Line [Total Bill],' + @cols + '  " & _
                                          " from (" & _
               " SELECT f.DOCNUM   , f.NUMATCARD  ,f.TAXDATE  ,f.DOCDATE  ,f.CARDNAME  ,f.JRNLMEMO  ,  b.[OcrCode3]  ,h.SLPNAME  ," & _
               "a.AcctCode + '' - '' + a.AcctName [plan_id] , e.[PrcCode] [U_AB_NONPROJECT], " & _
                "  sum(ISNULL( g.LineTotal,0)) [Line] , sum( round((ISNULL( g.LineTotal,0)) /(d.[OcrTotal] /e.[PrcAmount] ),3)) AS total " & _
                " FROM JDT1 b JOIN OACT a ON a.AcctCode = b.Account " & _
    " JOIN OJDT c on b.[TransId] = c.[TransId] " & _
    " left join OOCR d on d.OcrCode = b.OcrCode3 " & _
    " join OCR1 e on d.[OcrCode] = e.[OcrCode] " & _
    " JOIN ORPC f on c.[TransId] = f.[TransId]  JOIN RPC1 g on g.[DocEntry] = f.[DocEntry] and g.AcctCode = b.Account  " & _
    "  LEFT JOIN OSLP h on f.SLPCODE=h.SLPCODE " & _
    " where  b.OcrCode3 in (" & sQuotes & ") and a.GroupMask = 5  and isnull(c.Indicator,'''') <> ''CA'' and " & _
    "	b.Account >= ''" & sSplitG(0) & "'' and b.Account <= ''" & sSplitG(1) & "'' and c.transType IN (''19'') and month(c.refdate) >= " & sSplitM(0) & " and month(c.refdate) <= " & sSplitM(1) & " and year(c.refdate) = " & Now.Year & " " & _
    " and  e.[ValidFrom] <= ''" & sNow & "'' and (e.[ValidTo] >= ''" & sNow & "'' or  isnull(e.[ValidTo],'''') = '''')" & _
             "  group by a.AcctCode,a.AcctName,e.[PrcCode] , f.DOCNUM   , f.NUMATCARD  ,f.TAXDATE  ,f.DOCDATE  ,f.CARDNAME  ,f.JRNLMEMO  ,  b.[OcrCode3]  ,h.SLPNAME,g.LineNum   ) x pivot ( sum(total) " & _
             "   for U_AB_NONPROJECT in (' + @cols + ') " & _
           " ) p " & _
    "  union all SELECT cast(Number as nvarchar(30)) ''SAP Reference Number''  , REF2 ''Vendor Invoice Number'' , TAXDATE ''Vendor Invoice Date'' ,REFDATE ''SAP Posting Date'' " & _
        " ,''JE'' ''Vendor / Payee Name'',MEMO ''Journal Remark'' ,  [OcrCode3] ''Distribution Code'', '''' ''Purchasing Department'', plan_id [Expenses],  Line [Total Bill],' + @cols + '  " & _
        " from (" & _
               " SELECT c.Number   , b.REF2  ,b.TAXDATE  ,c.REFDATE  ,c.MEMO  ,  b.[OcrCode3]  , a.AcctCode + '' - '' + a.AcctName [plan_id] , e.[PrcCode] [U_AB_NONPROJECT], " & _
                "   sum( round((ISNULL(b.Debit,0) - ISNULL(b.Credit,0)) /(d.[OcrTotal] /e.[PrcAmount] ),3)) AS total , sum(ISNULL(b.Debit,0) - ISNULL(b.Credit,0)) [Line] " & _
                " FROM JDT1 b JOIN OACT a ON a.AcctCode = b.Account " & _
    " JOIN OJDT c on b.[TransId] = c.[TransId] " & _
    " left join OOCR d on d.OcrCode = b.OcrCode3 " & _
    " join OCR1 e on d.[OcrCode] = e.[OcrCode] " & _
    " where  b.OcrCode3 in (" & sQuotes & ") and a.GroupMask = 5  and isnull(c.Indicator,'''') <> ''CA'' and " & _
    "	b.Account >= ''" & sSplitG(0) & "'' and b.Account <= ''" & sSplitG(1) & "'' and c.transType IN (''30'',''1470000071'') and month(c.refdate) >= " & sSplitM(0) & " and month(c.refdate) <= " & sSplitM(1) & " and year(c.refdate) = " & Now.Year & " " & _
    " and  e.[ValidFrom] <= ''" & sNow & "'' and (e.[ValidTo] >= ''" & sNow & "'' or  isnull(e.[ValidTo],'''') = '''')" & _
    " group by c.Number   , b.REF2  ,b.TAXDATE  ,c.REFDATE  ,c.MEMO  ,  b.[OcrCode3]  , a.AcctCode , a.AcctName , e.[PrcCode] " & _
             "   ) x pivot ( sum(total) " & _
             "   for U_AB_NONPROJECT in (' + @cols + ') " & _
           " ) p ' as varchar(max))" & _
    " execute (@query)"


                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("SummaryReport show button click " & sSQL, sFuncName)
                                        oGridT2 = oForm.Items.Item("Item_23").Specific
                                        Try
                                            oForm.DataSources.DataTables.Add("MyDataTable1")
                                        Catch ex As Exception
                                        End Try
                                        oGridT2.DataTable = oForm.DataSources.DataTables.Item("MyDataTable1")
                                        oForm.DataSources.DataTables.Item(1).ExecuteQuery(sSQL)
                                        oGridT2.CollapseLevel = 1
                                        oGridT2.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
                                        Dim brow As Integer = oGridT2.DataTable.Rows.Count
                                        oGridT2.DataTable.Rows.Add(1)
                                        '' oForm.DataSources.DataTables.Item(1).SetValue(0, brow, "Total")
                                        '' oGridT2.DataTable.SetValue(8, brow, "Total")
                                        oGridT2.DataTable.SetValue(0, brow, "Total")
                                        Dim bcount As Double = 0
                                        For imjs As Integer = 9 To oGridT2.Columns.Count - 1
                                            bcount = 0
                                            For imjd As Integer = 0 To brow - 1
                                                If Not String.IsNullOrEmpty(oGridT2.DataTable.GetValue(imjs, imjd)) Then
                                                    bcount += oGridT2.DataTable.GetValue(imjs, imjd)
                                                End If
                                            Next
                                            oGridT2.DataTable.SetValue(imjs, brow, Math.Round(bcount))
                                        Next
                                        ''  oGridT2.CommonSetting.FixedColumnsCount = 2
                                        oGridT2.CommonSetting.SetRowBackColor(oGridT2.Rows.Count, 234)
                                        oGridT2.AutoResizeColumns()

                                        ''oGridT2.DataTable.SetValue(0, brow, "Total")
                                        ''Dim bcount As Double = 0
                                        ''For imjs As Integer = 9 To oGridT2.Columns.Count - 1
                                        ''    bcount = 0
                                        ''    For imjd As Integer = 0 To brow - 1
                                        ''        If Not String.IsNullOrEmpty(oGridT2.DataTable.GetValue(imjs, imjd)) Then
                                        ''            bcount += oGridT2.DataTable.GetValue(imjs, imjd)
                                        ''        End If
                                        ''    Next
                                        ''    oGridT2.DataTable.SetValue(imjs, brow, Math.Round(bcount))
                                        ''Next
                                        ''oGridT2.CommonSetting.FixedColumnsCount = 2
                                        ''oGridT2.CommonSetting.SetRowBackColor(oGridT2.Rows.Count, 234)
                                        ''oGridT2.AutoResizeColumns()


                                        ''oGridT2.DataTable = oForm.DataSources.DataTables.Item("MyDataTable1")
                                        ''oForm.DataSources.DataTables.Item(1).ExecuteQuery(sSQL)
                                        ''oGridT2.CollapseLevel = 1
                                        ''oGridT2.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
                                        ''oGridT2.AutoResizeColumns()

                                        oForm.Freeze(False)
                                        p_oSBOApplication.SetStatusBarMessage("Cost Allocation information successfully loaded .......!", SAPbouiCOM.BoMessageTime.bmt_Short, False)

                                    Catch ex As Exception
                                        oForm.Freeze(False)
                                        p_oSBOApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                        BubbleEvent = False
                                        sErrDesc = ex.Message
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                                        WriteToLogFile(Err.Description, sFuncName)
                                        ShowErr(sErrDesc)
                                    End Try
                                End If

                                If pVal.ItemUID = "31" Then


                                    Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                                    Dim sMonth As String = String.Empty
                                    Dim sSplitM() As String
                                    Dim sSplitD() As String
                                    Dim sSplitG() As String
                                    Dim sDistribution As String = String.Empty
                                    Dim sGLAccount As String = String.Empty
                                    Dim sSQL As String = String.Empty
                                    Dim oGridT1 As SAPbouiCOM.Grid = Nothing
                                    Dim oGridT2 As SAPbouiCOM.Grid = Nothing
                                    Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("Item_24").Specific
                                    Dim sNow As String = String.Empty
                                    Dim sIN As String = String.Empty
                                    Dim sQuotes As String = String.Empty
                                    Dim oRset As SAPbobsCOM.Recordset = Nothing
                                    Dim sOUF As String = String.Empty
                                    Dim sOUT As String = String.Empty
                                    Dim oDT As DataTable = Nothing

                                    Try

                                        oRset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        sFuncName = "Summary Report Show button click()"
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling HeaderValidation()", sFuncName)
                                        If CostAllocation_Validation(oForm, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                        If String.IsNullOrEmpty(p_Summaryreport) Then
                                            p_oSBOApplication.SetStatusBarMessage("Operating Unit should not be Empty ...... !", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                        oForm.Freeze(True)
                                        sMonth = oForm.Items.Item("Item_13").Specific.selected.description & "," & oForm.Items.Item("Item_14").Specific.selected.description
                                        ''  sDistribution = oForm.Items.Item("Item_5").Specific.value & "," & oForm.Items.Item("Item_8").Specific.value
                                        sGLAccount = oForm.Items.Item("Item_10").Specific.value & "," & oForm.Items.Item("Item_12").Specific.value
                                        sSplitM = sMonth.Split(",")
                                        '' sSplitD = sDistribution.Split(",")
                                        sSplitG = sGLAccount.Split(",")
                                        ''sOUF = oForm.Items.Item("Item_22").Specific.String
                                        ''sOUT = oForm.Items.Item("Item_26").Specific.String

                                        ''sSQL = "DECLARE  " & _
                                        ''   "  @string varchar(100), @string1 varchar(100), @string2 varchar(100), @string3 varchar(max), @string4 varchar(max) " & _
                                        ''   "  SET @string = '" & sSplitD(0) & "' SET @string1 = '" & sSplitD(1) & "' SET @string2 = '" & sSplitD(0) & "' " & _
                                        ''   " WHILE PATINDEX('%[^a-z]%',@string2) > 0 SET @string2 = STUFF(@string2,PATINDEX('%[^a-z]%',@string2),1,'') " & _
                                        ''   " WHILE PATINDEX('%[^0-9]%',@string) <> 0     SET @string = STUFF(@string,PATINDEX('%[^0-9]%',@string),1,'') " & _
                                        ''   " WHILE PATINDEX('%[^0-9]%',@string1) <> 0     SET @string1 = STUFF(@string1,PATINDEX('%[^0-9]%',@string1),1,'') " & _
                                        ''   " WHILE cast(@string as numeric ) <= cast(@string1 as numeric ) " & _
                                        ''   " begin " & _
                                        ''   " SET @string3 =  isnull(@string3,'')  + '[' + @string2 + @string + '],' " & _
                                        ''   " set @string = cast(@string as numeric ) + 1 " & _
                                        ''   " end " & _
                                        ''   " set @string3 = replace( replace(@string3,'[',''''),']','''') " & _
                                        ''   " set @string3 = left(@string3, len(@string3) -1) " & _
                                        ''     " set @string4 = replace(@string3,'''','''''') " & _
                                        ''   " select @string3 [Ouput] , @string4 [quotes] "
                                        ''oRset.DoQuery(sSQL)
                                        ''sIN = oRset.Fields.Item("Ouput").Value
                                        ''sQuotes = oRset.Fields.Item("quotes").Value
                                        sIN = p_Dimensionrules
                                        sQuotes = Replace(p_Dimensionrules, "'", "''")
                                        oGridT2 = oForm.Items.Item("Item_23").Specific
                                        sNow = CStr(Now.Year) & sSplitM(0).PadLeft(2, "0"c) & "01"


                                        sSQL = "DECLARE @cols AS NVARCHAR(MAX), @cols3 AS NVARCHAR(MAX),  @query  AS VARCHAR(max), @query1  AS VARCHAR(max), @cols1 as nvarchar(max) " & _
                                              "select @cols3 = STUFF((SELECT distinct  ',' + QUOTENAME(cast(T1.PrcCode as nvarchar(100)))   + ' [Column' + cast(ROW_NUMBER() OVER(ORDER BY T1.PrcCode asc) as varchar) + ']' " & _
                   " from OOCR T0  INNER JOIN OCR1 T1 ON T0.[OcrCode] = T1.[OcrCode]  where T0.OcrCode in (" & sIN & ") and T1.[ValidFrom] <= '" & sNow & "' and  (T1.[ValidTo] >= '" & sNow & "' or  isnull(T1.[ValidTo],'') = '')" & _
                   " and T1.PrcCode in (" & p_Summaryreport & ") " & _
           " FOR XML PATH(''), TYPE " & _
           " ).value('.', 'NVARCHAR(MAX)') " & _
       "  ,1,1,'') " & _
    "select @cols1 = STUFF((SELECT '+ isnull(' + QUOTENAME(isnull(T1.PrcCode,0)) + ',0)' " & _
                   " from OOCR T0  INNER JOIN OCR1 T1 ON T0.[OcrCode] = T1.[OcrCode]  where T0.OcrCode in (" & sIN & ") and T1.[ValidFrom] <= '" & sNow & "' and  (T1.[ValidTo] >= '" & sNow & "' or  isnull(T1.[ValidTo],'') = '')" & _
                    " and T1.PrcCode in (" & p_Summaryreport & ") " & _
           " FOR XML PATH(''), TYPE " & _
           " ).value('.', 'NVARCHAR(MAX)') " & _
       "  ,1,1,'') " & _
       " select @cols = STUFF((SELECT distinct  ',' + QUOTENAME(cast(T1.PrcCode as nvarchar(100)))   " & _
                   " from OOCR T0  INNER JOIN OCR1 T1 ON T0.[OcrCode] = T1.[OcrCode]  where T0.OcrCode in (" & sIN & ") and T1.[ValidFrom] <= '" & sNow & "' and  (T1.[ValidTo] >= '" & sNow & "' or  isnull(T1.[ValidTo],'') = '')" & _
                    " and T1.PrcCode in (" & p_Summaryreport & ") " & _
           " FOR XML PATH(''), TYPE " & _
           " ).value('.', 'NVARCHAR(MAX)') " & _
       "  ,1,1,'') " & _
       " set @query = cast('SELECT SeriesName , ''PU '' + cast(DOCNUM as varchar(30)) ''SAP Reference Number'' , PDocNum ''Payment Doc Num'', PDocDate ''Payment Doc Date'' ,  NUMATCARD ''Vendor Invoice Number'',TAXDATE ''Vendor Invoice Date'',DOCDATE ''SAP Posting Date'',  " & _
    " CARDNAME ''Vendor / Payee Name'',JRNLMEMO ''Journal Remark'',  [OcrCode3] ''Distribution Code'', SLPNAME ''Purchasing Department'', Creator [Document Creator] , " & _
    "plan_id [Expenses], Line [Total Bill],' + @cols + '  " & _
                                          " from (" & _
               " SELECT i.SeriesName, f.DOCNUM   , f.NUMATCARD  ,f.TAXDATE  ,f.DOCDATE  ,f.CARDNAME  ,f.JRNLMEMO  ,  b.[OcrCode3]  ,h.SLPNAME  ," & _
               "a.AcctCode + '' - '' + a.AcctName [plan_id] , e.[PrcCode] [U_AB_NONPROJECT], l.DocNum [PDocNum] , l.DocDate [PDocDate] , g.LineNum , " & _
                "  case when g.BaseType = 18 then -sum(isnull(g.LineTotal,0)) else sum(isnull(g.LineTotal,0)) end [Line] , " & _
                " case when g.BaseType = 18 then -sum( round(isnull(g.LineTotal,0) /(d.[OcrTotal] /e.[PrcAmount] ),3)) else sum( round(isnull(g.LineTotal,0) /(d.[OcrTotal] /e.[PrcAmount] ),3)) end AS total , j.[U_NAME] +'', '' + j.[USER_CODE][Creator] " & _
                " FROM JDT1 b JOIN OACT a ON a.AcctCode = b.Account " & _
    " JOIN OJDT c on b.[TransId] = c.[TransId] " & _
    " left join OOCR d on d.OcrCode = b.OcrCode3 " & _
    " join OCR1 e on d.[OcrCode] = e.[OcrCode] " & _
    " JOIN OPCH f on c.[TransId] = f.[TransId]  JOIN PCH1 g on g.[DocEntry] = f.[DocEntry] and g.AcctCode = b.Account and  g.[OcrCode3] =  b.[OcrCode3] " & _
    "  LEFT JOIN OSLP h on f.SLPCODE=h.SLPCODE " & _
    "  LEFT JOIN NNM1 i on i.Series=f.Series " & _
    "   LEFT JOIN OUSR j on  f.[usersign] = j.[USERID] " & _
    " left join   [VPM2] k on f.DocEntry = k.DocEntry left JOIN OVPM l ON k.[DocNum] = l.[DocEntry] " & _
    " where  b.OcrCode3 in (" & sQuotes & ") and a.GroupMask = 5  and isnull(c.Indicator,'''') <> ''CA'' and " & _
    "	b.Account >= ''" & sSplitG(0) & "'' and b.Account <= ''" & sSplitG(1) & "'' and c.transType IN (''18'') and month(c.refdate) >= " & sSplitM(0) & " and month(c.refdate) <= " & sSplitM(1) & " and year(c.refdate) = " & Now.Year & " " & _
    " and  e.[ValidFrom] <= ''" & sNow & "'' and (e.[ValidTo] >= ''" & sNow & "'' or  isnull(e.[ValidTo],'''') = '''')" & _
     " and e.PrcCode not in (SELECT T0.[U_PrcCode] FROM [dbo].[@AB_COSTALLOCATION]  T0 WHERE T0.[U_Transid] = c.TransId and T0.U_OcrCode3 = b.OcrCode3) and (case when isnull(l.DocNum,'''') = '''' then ''N'' else l.Canceled end)=''N'' " & _
             "  group by a.AcctCode,a.AcctName,e.[PrcCode] , f.DOCNUM   , f.NUMATCARD  ,f.TAXDATE  ,f.DOCDATE  ,f.CARDNAME  ,f.JRNLMEMO  ,  b.[OcrCode3]  ,h.SLPNAME , g.LineNum, i.SeriesName ,j.[U_NAME] , j.[USER_CODE], l.DocNum , l.DocDate , g.BaseType  ) x pivot ( sum(total) " & _
             "   for U_AB_NONPROJECT in (' + @cols + ') " & _
           " ) p  " & _
               " union all SELECT SeriesName, ''PC '' + cast(DOCNUM as varchar(30)) ''SAP Reference Number'' , '''' ''Payment Doc Num'', NULL ''Payment Doc Date'',  NUMATCARD ''Vendor Invoice Number'',TAXDATE ''Vendor Invoice Date'',DOCDATE ''SAP Posting Date'',  " & _
    " CARDNAME ''Vendor / Payee Name'',JRNLMEMO ''Journal Remark'',  [OcrCode3] ''Distribution Code'', SLPNAME ''Purchasing Department'', Creator [Document Creator], " & _
    "plan_id [Expenses], Line [Total Bill],' + @cols + '  " & _
                                          " from (" & _
               " SELECT i.SeriesName, f.DOCNUM   , f.NUMATCARD  ,f.TAXDATE  ,f.DOCDATE  ,f.CARDNAME  ,f.JRNLMEMO  ,  b.[OcrCode3]  ,h.SLPNAME  , j.[U_NAME] +'', '' + j.[USER_CODE][Creator] ," & _
               "a.AcctCode + '' - '' + a.AcctName [plan_id] , e.[PrcCode] [U_AB_NONPROJECT], g.LineNum, " & _
                " case when g.BaseType = ''19'' then   sum(ISNULL( g.LineTotal,0))  else  - sum(ISNULL( g.LineTotal,0))  end [Line] , " & _
                " case when g.BaseType = ''19'' then  sum( round((isnull( g.LineTotal,0)) /(d.[OcrTotal] /e.[PrcAmount] ),3)) else - sum( round((isnull( g.LineTotal,0)) /(d.[OcrTotal] /e.[PrcAmount] ),3)) end AS total " & _
                " FROM JDT1 b JOIN OACT a ON a.AcctCode = b.Account " & _
    " JOIN OJDT c on b.[TransId] = c.[TransId] " & _
    " left join OOCR d on d.OcrCode = b.OcrCode3 " & _
    " join OCR1 e on d.[OcrCode] = e.[OcrCode] " & _
    " JOIN ORPC f on c.[TransId] = f.[TransId]  JOIN RPC1 g on g.[DocEntry] = f.[DocEntry] and g.AcctCode = b.Account  and  g.[OcrCode3] =  b.[OcrCode3] " & _
    "  LEFT JOIN OSLP h on f.SLPCODE=h.SLPCODE " & _
    "  LEFT JOIN NNM1 i on i.Series=f.Series " & _
    "   LEFT JOIN OUSR j on  f.[usersign] = j.[USERID] " & _
    " where  b.OcrCode3 in (" & sQuotes & ") and a.GroupMask = 5  and isnull(c.Indicator,'''') <> ''CA'' and " & _
    "	b.Account >= ''" & sSplitG(0) & "'' and b.Account <= ''" & sSplitG(1) & "'' and c.transType IN (''19'') and month(c.refdate) >= " & sSplitM(0) & " and month(c.refdate) <= " & sSplitM(1) & " and year(c.refdate) = " & Now.Year & " " & _
    " and  e.[ValidFrom] <= ''" & sNow & "'' and (e.[ValidTo] >= ''" & sNow & "'' or  isnull(e.[ValidTo],'''') = '''')" & _
     " and e.PrcCode not in (SELECT T0.[U_PrcCode] FROM [dbo].[@AB_COSTALLOCATION]  T0 WHERE T0.[U_Transid] = c.TransId and T0.U_OcrCode3 = b.OcrCode3) " & _
             "  group by a.AcctCode,a.AcctName,e.[PrcCode] , f.DOCNUM   , f.NUMATCARD  ,f.TAXDATE  ,f.DOCDATE  ,f.CARDNAME  ,f.JRNLMEMO  ,  b.[OcrCode3]  ,h.SLPNAME , g.LineNum , i.SeriesName ,j.[U_NAME] , j.[USER_CODE] ,g.BaseType ) x pivot ( sum(total) " & _
             "   for U_AB_NONPROJECT in (' + @cols + ') " & _
           " ) p ' as nvarchar(max)) " & _
  " set @query1 = cast( ' union all SELECT SeriesName, ''JE '' + cast(Number as varchar(30)) ''SAP Reference Number''  , '''' ''Payment Doc Num'', NULL ''Payment Doc Date'' ,REF2 ''Vendor Invoice Number'' , TAXDATE ''Vendor Invoice Date'' ,REFDATE ''SAP Posting Date'' " & _
        " ,''JE'' ''Vendor / Payee Name'',MEMO ''Journal Remark'' ,  [OcrCode3] ''Distribution Code'', '''' ''Purchasing Department'', Creator [Document Creator], plan_id [Expenses],  Line [Total Bill],' + @cols + '  " & _
        " from (" & _
               " SELECT i.SeriesName, c.Number   , b.REF2  ,b.TAXDATE  ,c.REFDATE  ,c.MEMO  ,  b.[OcrCode3]  , a.AcctCode + '' - '' + a.AcctName [plan_id] , e.[PrcCode] [U_AB_NONPROJECT] , j.[U_NAME] +'', '' + j.[USER_CODE][Creator] ," & _
                "   sum( round((ISNULL(b.Debit,0) - ISNULL(b.Credit,0)) /(d.[OcrTotal] /e.[PrcAmount] ),3)) AS total , sum(ISNULL(b.Debit,0) - ISNULL(b.Credit,0)) [Line] " & _
                " FROM JDT1 b JOIN OACT a ON a.AcctCode = b.Account " & _
    " JOIN OJDT c on b.[TransId] = c.[TransId] " & _
    " left join OOCR d on d.OcrCode = b.OcrCode3 " & _
    " join OCR1 e on d.[OcrCode] = e.[OcrCode] " & _
    "  LEFT JOIN NNM1 i on i.Series=c.Series " & _
    "  LEFT JOIN OUSR j on  c.[usersign] = j.[USERID]" & _
    " where  b.OcrCode3 in (" & sQuotes & ") and a.GroupMask = 5  and isnull(c.Indicator,'''') <> ''CA'' and " & _
    "	b.Account >= ''" & sSplitG(0) & "'' and b.Account <= ''" & sSplitG(1) & "'' and c.transType IN (''30'',''1470000071'') and month(c.refdate) >= " & sSplitM(0) & " and month(c.refdate) <= " & sSplitM(1) & " and year(c.refdate) = " & Now.Year & " " & _
    " and  e.[ValidFrom] <= ''" & sNow & "'' and (e.[ValidTo] >= ''" & sNow & "'' or  isnull(e.[ValidTo],'''') = '''')" & _
     " and e.PrcCode not in (SELECT T0.[U_PrcCode] FROM [dbo].[@AB_COSTALLOCATION]  T0 WHERE T0.[U_Transid] = c.TransId and T0.U_OcrCode3 = b.OcrCode3) " & _
    " group by c.Number   , b.REF2  ,b.TAXDATE  ,c.REFDATE  ,c.MEMO  ,  b.[OcrCode3]  , a.AcctCode , a.AcctName , e.[PrcCode],i.SeriesName,j.[U_NAME] , j.[USER_CODE] " & _
             "   ) x pivot ( sum(total) " & _
             "   for U_AB_NONPROJECT in (' + @cols + ') " & _
           " ) p " & _
           " UNION ALL SELECT  SeriesName, ''PS '' + cast(DOCNUM as varchar(30))''SAP Reference Number'' , '''' ''Payment Doc Num'', NULL ''Payment Doc Date'' ,'''' ''Vendor Invoice Number'',TAXDATE ''Vendor Invoice Date'',DOCDATE ''SAP Posting Date'',  " & _
    " Address ''Vendor / Payee Name'',JRNLMEMO ''Journal Remark'',  [OcrCode3] ''Distribution Code'', '''' ''Purchasing Department'', Creator [Document Creator], " & _
    "plan_id [Expenses], Line [Total Bill],' + @cols + '  " & _
                                          " from (" & _
               " SELECT i.SeriesName,f.DOCNUM   ,f.TAXDATE  ,f.DOCDATE   ,f.JRNLMEMO  ,  b.[OcrCode3]   ," & _
               "a.AcctCode + '' - '' + a.AcctName [plan_id] , e.[PrcCode] [U_AB_NONPROJECT],  f.[Address] , g.LineId,  j.[U_NAME] +'', '' + j.[USER_CODE][Creator] ," & _
                "  sum(isnull(g.SumApplied,0)) [Line] , sum( round(isnull(g.SumApplied,0) /(d.[OcrTotal] /e.[PrcAmount] ),3)) AS total " & _
                " FROM JDT1 b JOIN OACT a ON a.AcctCode = b.Account " & _
    " JOIN OJDT c on b.[TransId] = c.[TransId] " & _
    " left join OOCR d on d.OcrCode = b.OcrCode3 " & _
    " join OCR1 e on d.[OcrCode] = e.[OcrCode] " & _
    " JOIN OVPM f on c.[TransId] = f.[TransId]  JOIN VPM4 g on g.[DocNum] = f.[DocNum] and g.AcctCode = b.Account and  g.[OcrCode3] =  b.[OcrCode3] " & _
     "  LEFT JOIN NNM1 i on i.Series=f.Series " & _
      "   LEFT JOIN OUSR j on  f.[usersign] = j.[USERID] " & _
     " where  b.OcrCode3 in (" & sQuotes & ") and a.GroupMask = 5  and isnull(c.Indicator,'''') <> ''CA'' and " & _
    "	b.Account >= ''" & sSplitG(0) & "'' and b.Account <= ''" & sSplitG(1) & "'' and c.transType IN (''46'') and month(c.refdate) >= " & sSplitM(0) & " and month(c.refdate) <= " & sSplitM(1) & " and year(c.refdate) = " & Now.Year & " " & _
    " and  e.[ValidFrom] <= ''" & sNow & "'' and (e.[ValidTo] >= ''" & sNow & "'' or  isnull(e.[ValidTo],'''') = '''')" & _
     " and e.PrcCode not in (SELECT T0.[U_PrcCode] FROM [dbo].[@AB_COSTALLOCATION]  T0 WHERE T0.[U_Transid] = c.TransId and T0.U_OcrCode3 = b.OcrCode3) " & _
     "  and f.DocType = ''A'' " & _
             "  group by a.AcctCode,a.AcctName,e.[PrcCode] , f.DOCNUM ,f.TAXDATE  ,f.DOCDATE  ,f.JRNLMEMO  ,  b.[OcrCode3]  , g.LineId , f.[Address],i.SeriesName ,j.[U_NAME] , j.[USER_CODE] ) x pivot ( sum(total) " & _
             "   for U_AB_NONPROJECT in (' + @cols + ') " & _
           " ) p " & _
           " UNION ALL SELECT  SeriesName, ''PS '' + cast(DOCNUM as varchar(30))''SAP Reference Number'' , '''' ''Payment Doc Num'', NULL ''Payment Doc Date'' ,'''' ''Vendor Invoice Number'',TAXDATE ''Vendor Invoice Date'',DOCDATE ''SAP Posting Date'',  " & _
    " Address ''Vendor / Payee Name'',JRNLMEMO ''Journal Remark'',  [OcrCode3] ''Distribution Code'', '''' ''Purchasing Department'', Creator [Document Creator], " & _
    "plan_id [Expenses], Line [Total Bill],' + @cols + '  " & _
                                          " from (" & _
               " SELECT i.SeriesName,f.DOCNUM   ,f.CancelDate [TAXDATE]  ,f.CancelDate [DOCDATE]   ,f.JRNLMEMO  ,  b.[OcrCode3]   ," & _
               "a.AcctCode + '' - '' + a.AcctName [plan_id] , e.[PrcCode] [U_AB_NONPROJECT],  f.[Address] , g.LineId ,  j.[U_NAME] +'', '' + j.[USER_CODE][Creator] ," & _
                "  -sum(isnull(g.SumApplied,0)) [Line] , -sum( round(isnull(g.SumApplied,0) /(d.[OcrTotal] /e.[PrcAmount] ),3)) AS total " & _
                " FROM JDT1 b JOIN OACT a ON a.AcctCode = b.Account " & _
    " JOIN OJDT c on b.[TransId] = c.[TransId] " & _
    " left join OOCR d on d.OcrCode = b.OcrCode3 " & _
    " join OCR1 e on d.[OcrCode] = e.[OcrCode] " & _
    " JOIN OVPM f on c.[TransId] = f.[TransId]  JOIN VPM4 g on g.[DocNum] = f.[DocNum] and g.AcctCode = b.Account  and  g.[OcrCode3] =  b.[OcrCode3] " & _
     "  LEFT JOIN NNM1 i on i.Series=f.Series " & _
      "   LEFT JOIN OUSR j on  f.[usersign] = j.[USERID] " & _
     " where  b.OcrCode3 in (" & sQuotes & ") and a.GroupMask = 5  and isnull(c.Indicator,'''') <> ''CA'' and " & _
    "	b.Account >= ''" & sSplitG(0) & "'' and b.Account <= ''" & sSplitG(1) & "'' and c.transType IN (''46'') and month(f.CancelDate) >= " & sSplitM(0) & " and month(f.CancelDate) <= " & sSplitM(1) & " and year(f.CancelDate) = " & Now.Year & " " & _
    " and  e.[ValidFrom] <= ''" & sNow & "'' and (e.[ValidTo] >= ''" & sNow & "'' or  isnull(e.[ValidTo],'''') = '''')" & _
     " and e.PrcCode not in (SELECT T0.[U_PrcCode] FROM [dbo].[@AB_COSTALLOCATION]  T0 WHERE T0.[U_Transid] = c.TransId and T0.U_OcrCode3 = b.OcrCode3) " & _
     "  and f.DocType = ''A'' and f.Canceled = ''Y''" & _
             "  group by a.AcctCode,a.AcctName,e.[PrcCode] , f.DOCNUM ,f.CancelDate,f.JRNLMEMO  ,  b.[OcrCode3]  , g.LineId , f.[Address],i.SeriesName ,j.[U_NAME] , j.[USER_CODE]  ) x pivot ( sum(total) " & _
             "   for U_AB_NONPROJECT in (' + @cols + ') " & _
           " ) p ' as nvarchar(max)) " & _
  " execute (@query + @query1)"

                                        ' ''" set @query = cast('SELECT  cast(DOCNUM as varchar(30)) ''Number'' , NUMATCARD ''InvoiceNumber'',TAXDATE ''InvoiceDate'',DOCDATE ''PostingDate'',  " & _
                                        '                                    '                                    '" CARDNAME ''VendorName'',JRNLMEMO ''JournalRemark'',  [OcrCode3] ''DistributionCode'', SLPNAME ''PurchasingDepartment'', " & _
                                        '                                    '                                    '"plan_id [Expenses], Line [TotalBill],' + @cols3 + '  " & _
                                        '                                    '                                    '                                      " from (" & _
                                        '                                    '                                    '           " SELECT f.DOCNUM   , f.NUMATCARD  ,f.TAXDATE  ,f.DOCDATE  ,f.CARDNAME  ,f.JRNLMEMO  ,  b.[OcrCode3]  ,h.SLPNAME  ," & _
                                        '                                    '                                    '           "a.AcctCode + '' - '' + a.AcctName [plan_id] , e.[PrcCode] [U_AB_NONPROJECT], " & _
                                        '                                    '                                    '            "  sum(ISNULL( g.LineTotal,0)) [Line] , sum( round((isnull( g.LineTotal,0)) /(d.[OcrTotal] /e.[PrcAmount] ),2)) AS total  " & _
                                        '                                    '                                    '            " FROM JDT1 b JOIN OACT a ON a.AcctCode = b.Account " & _
                                        '                                    '                                    '" JOIN OJDT c on b.[TransId] = c.[TransId] " & _
                                        '                                    '                                    '" left join OOCR d on d.OcrCode = b.OcrCode3 " & _
                                        '                                    '                                    '" join OCR1 e on d.[OcrCode] = e.[OcrCode] " & _
                                        '                                    '                                    '" JOIN OPCH f on c.[TransId] = f.[TransId]  JOIN PCH1 g on g.[DocEntry] = f.[DocEntry] and g.AcctCode = b.Account  and  g.[OcrCode3] =  b.[OcrCode3] " & _
                                        '                                    '                                    '"  LEFT JOIN OSLP h on f.SLPCODE=h.SLPCODE " & _
                                        '                                    '                                    '" where  b.OcrCode3 in (" & sQuotes & ") and a.GroupMask = 5  and isnull(c.Indicator,'''') <> ''CA'' and " & _
                                        '                                    '                                    '"	b.Account >= ''" & sSplitG(0) & "'' and b.Account <= ''" & sSplitG(1) & "'' and c.transType IN (''18'') and month(c.refdate) >= " & sSplitM(0) & " and month(c.refdate) <= " & sSplitM(1) & " and year(c.refdate) = " & Now.Year & " " & _
                                        '                                    '                                    '" and  e.[ValidFrom] <= ''" & sNow & "'' and (e.[ValidTo] >= ''" & sNow & "'' or  isnull(e.[ValidTo],'''') = '''')" & _
                                        '                                    '                                    '         "  group by a.AcctCode,a.AcctName,e.[PrcCode] , f.DOCNUM   , f.NUMATCARD  ,f.TAXDATE  ,f.DOCDATE  ,f.CARDNAME  ,f.JRNLMEMO  ,  b.[OcrCode3]  ,h.SLPNAME, g.LineNum   ) x pivot ( sum(total) " & _
                                        '                                    '                                    '         "   for U_AB_NONPROJECT in (' + @cols + ') " & _
                                        '                                    '                                    '       " ) p  " & _
                                        '                                    '                                    '           " union all SELECT  cast(DOCNUM as varchar(30)) ''Number'' , NUMATCARD ''InvoiceNumber'',TAXDATE ''InvoiceDate'',DOCDATE ''PostingDate'',  " & _
                                        '                                    '                                    '" CARDNAME ''VendorName'',JRNLMEMO ''JournalRemark'',  [OcrCode3] ''DistributionCode'', SLPNAME ''PurchasingDepartment'', " & _
                                        '                                    '                                    '"plan_id [Expenses], Line [TotalBill],' + @cols3 + '  " & _
                                        '                                    '                                    '                                      " from (" & _
                                        '                                    '                                    '           " SELECT f.DOCNUM   , f.NUMATCARD  ,f.TAXDATE  ,f.DOCDATE  ,f.CARDNAME  ,f.JRNLMEMO  ,  b.[OcrCode3]  ,h.SLPNAME  ," & _
                                        '                                    '                                    '           "a.AcctCode + '' - '' + a.AcctName [plan_id] , e.[PrcCode] [U_AB_NONPROJECT], " & _
                                        '                                    '                                    '            "  sum(ISNULL( g.LineTotal,0)) [Line] , sum( round((ISNULL( g.LineTotal,0)) /(d.[OcrTotal] /e.[PrcAmount] ),2)) AS total " & _
                                        '                                    '                                    '            " FROM JDT1 b JOIN OACT a ON a.AcctCode = b.Account " & _
                                        '                                    '                                    '" JOIN OJDT c on b.[TransId] = c.[TransId] " & _
                                        '                                    '                                    '" left join OOCR d on d.OcrCode = b.OcrCode3 " & _
                                        '                                    '                                    '" join OCR1 e on d.[OcrCode] = e.[OcrCode] " & _
                                        '                                    '                                    '" JOIN ORPC f on c.[TransId] = f.[TransId]  JOIN RPC1 g on g.[DocEntry] = f.[DocEntry] and g.AcctCode = b.Account  and  g.[OcrCode3] =  b.[OcrCode3]" & _
                                        '                                    '                                    '"  LEFT JOIN OSLP h on f.SLPCODE=h.SLPCODE " & _
                                        '                                    '                                    '" where  b.OcrCode3 in (" & sQuotes & ") and a.GroupMask = 5  and isnull(c.Indicator,'''') <> ''CA'' and " & _
                                        '                                    '                                    '"	b.Account >= ''" & sSplitG(0) & "'' and b.Account <= ''" & sSplitG(1) & "'' and c.transType IN (''19'') and month(c.refdate) >= " & sSplitM(0) & " and month(c.refdate) <= " & sSplitM(1) & " and year(c.refdate) = " & Now.Year & " " & _
                                        '                                    '                                    '" and  e.[ValidFrom] <= ''" & sNow & "'' and (e.[ValidTo] >= ''" & sNow & "'' or  isnull(e.[ValidTo],'''') = '''')" & _
                                        '                                    '                                    '         "  group by a.AcctCode,a.AcctName,e.[PrcCode] , f.DOCNUM   , f.NUMATCARD  ,f.TAXDATE  ,f.DOCDATE  ,f.CARDNAME  ,f.JRNLMEMO  ,  b.[OcrCode3]  ,h.SLPNAME,g.LineNum   ) x pivot ( sum(total) " & _
                                        '                                    '                                    '         "   for U_AB_NONPROJECT in (' + @cols + ') " & _
                                        '                                    '                                    '       " ) p " & _
                                        '                                    '                                    '"  union all SELECT cast(Number as varchar(30)) ''Number''  , REF2 ''InvoiceNumber'' , TAXDATE ''InvoiceDate'' ,REFDATE ''PostingDate'' " & _
                                        '                                    '                                    '    " ,''JE'' ''VendorName'',MEMO ''JournalRemark'' ,  [OcrCode3] ''DistributionCode'', '''' ''PurchasingDepartment'', plan_id [Expenses],  Line [TotalBill],' + @cols3 + '  " & _
                                        '                                    '                                    '    " from (" & _
                                        '                                    '                                    '           " SELECT c.Number   , b.REF2  ,b.TAXDATE  ,c.REFDATE  ,c.MEMO  ,  b.[OcrCode3]  , a.AcctCode + '' - '' + a.AcctName [plan_id] , e.[PrcCode] [U_AB_NONPROJECT], " & _
                                        '                                    '                                    '            "   sum( round((ISNULL(b.Debit,0) - ISNULL(b.Credit,0)) /(d.[OcrTotal] /e.[PrcAmount] ),2)) AS total , sum(ISNULL(b.Debit,0) - ISNULL(b.Credit,0)) [Line] " & _
                                        '                                    '                                    '            " FROM JDT1 b JOIN OACT a ON a.AcctCode = b.Account " & _
                                        '                                    '                                    '" JOIN OJDT c on b.[TransId] = c.[TransId] " & _
                                        '                                    '                                    '" left join OOCR d on d.OcrCode = b.OcrCode3 " & _
                                        '                                    '                                    '" join OCR1 e on d.[OcrCode] = e.[OcrCode] " & _
                                        '                                    '                                    '" where  b.OcrCode3 in (" & sQuotes & ") and a.GroupMask = 5  and isnull(c.Indicator,'''') <> ''CA'' and " & _
                                        '                                    '                                    '"	b.Account >= ''" & sSplitG(0) & "'' and b.Account <= ''" & sSplitG(1) & "'' and c.transType IN (''30'',''1470000071'') and month(c.refdate) >= " & sSplitM(0) & " and month(c.refdate) <= " & sSplitM(1) & " and year(c.refdate) = " & Now.Year & " " & _
                                        '                                    '                                    '" and  e.[ValidFrom] <= ''" & sNow & "'' and (e.[ValidTo] >= ''" & sNow & "'' or  isnull(e.[ValidTo],'''') = '''')" & _
                                        '                                    '                                    '" group by c.Number   , b.REF2  ,b.TAXDATE  ,c.REFDATE  ,c.MEMO  ,  b.[OcrCode3]  , a.AcctCode , a.AcctName , e.[PrcCode] " & _
                                        '                                    '                                    '         "   ) x pivot ( sum(total) " & _
                                        '                                    '                                    '         "   for U_AB_NONPROJECT in (' + @cols + ') " & _
                                        '                                    '                                    '       " ) p ' as varchar(max))" & _
                                        '                                    '                                    '" execute (@query)"


                                        oRset.DoQuery(sSQL)
                                        oDT = New DataTable
                                        oDT = ConvertRecordsetToDataTable(oRset, sErrDesc)
                                        Dim oDS As DataSet = Nothing
                                        oDS = New DataSet
                                        oDT.TableName = "Summary"
                                        oDS.Tables.Add(oDT)
                                        PrintCalling_Summary(oDS)
                                        oForm.Freeze(False)

                                    Catch ex As Exception
                                        p_oSBOApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                        BubbleEvent = False
                                        Exit Sub
                                    End Try


                                End If

                                If pVal.ItemUID = "Item_27" Then

                                    Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                                    Dim sMonth As String = String.Empty
                                    Dim sSplitM() As String
                                    Dim sSplitD() As String
                                    Dim sSplitG() As String
                                    Dim sDistribution As String = String.Empty
                                    Dim sGLAccount As String = String.Empty
                                    Dim sSQL As String = String.Empty
                                    Dim oGridT1 As SAPbouiCOM.Grid = Nothing
                                    Dim oGridT2 As SAPbouiCOM.Grid = Nothing
                                    Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("Item_24").Specific
                                    Dim sNow As String = String.Empty
                                    Dim sIN As String = String.Empty
                                    Dim sQuotes As String = String.Empty
                                    Dim oRset As SAPbobsCOM.Recordset = Nothing
                                    Dim sOU As String = String.Empty
                                    Dim sOUT As String = String.Empty

                                    Try

                                        oRset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        sFuncName = "Summary Report Show button click()"
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling HeaderValidation()", sFuncName)
                                        p_oSBOApplication.SetStatusBarMessage("Information Starting to Load on the Summary Report .......!", SAPbouiCOM.BoMessageTime.bmt_Short, False)

                                        If CostAllocation_Validation(oForm, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                        If String.IsNullOrEmpty(p_Summaryreport) Then
                                            p_oSBOApplication.SetStatusBarMessage("Operating Unit should not be Empty ...... !", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                        oForm.Freeze(True)
                                        sMonth = oForm.Items.Item("Item_13").Specific.selected.description & "," & oForm.Items.Item("Item_14").Specific.selected.description
                                        '' sDistribution = oForm.Items.Item("Item_5").Specific.value & "," & oForm.Items.Item("Item_8").Specific.value
                                        sGLAccount = oForm.Items.Item("Item_10").Specific.value & "," & oForm.Items.Item("Item_12").Specific.value
                                        sSplitM = sMonth.Split(",")
                                        '' sSplitD = sDistribution.Split(",")
                                        sSplitG = sGLAccount.Split(",")
                                        ''sOU = p_Summaryreport
                                        ''sOUF = oForm.Items.Item("Item_22").Specific.String
                                        ''sOUT = oForm.Items.Item("Item_26").Specific.String

                                        ''oForm.Items.Item("Item_2").Specific.String = sOUF
                                        ''oForm.Items.Item("Item_15").Specific.String = sOUT

                                        ''sSQL = "DECLARE  " & _
                                        ''   "  @string varchar(100), @string1 varchar(100), @string2 varchar(100), @string3 varchar(max), @string4 varchar(max) " & _
                                        ''   "  SET @string = '" & sSplitD(0) & "' SET @string1 = '" & sSplitD(1) & "' SET @string2 = '" & sSplitD(0) & "' " & _
                                        ''   " WHILE PATINDEX('%[^a-z]%',@string2) > 0 SET @string2 = STUFF(@string2,PATINDEX('%[^a-z]%',@string2),1,'') " & _
                                        ''   " WHILE PATINDEX('%[^0-9]%',@string) <> 0     SET @string = STUFF(@string,PATINDEX('%[^0-9]%',@string),1,'') " & _
                                        ''   " WHILE PATINDEX('%[^0-9]%',@string1) <> 0     SET @string1 = STUFF(@string1,PATINDEX('%[^0-9]%',@string1),1,'') " & _
                                        ''   " WHILE cast(@string as numeric ) <= cast(@string1 as numeric ) " & _
                                        ''   " begin " & _
                                        ''   " SET @string3 =  isnull(@string3,'')  + '[' + @string2 + @string + '],' " & _
                                        ''   " set @string = cast(@string as numeric ) + 1 " & _
                                        ''   " end " & _
                                        ''   " set @string3 = replace( replace(@string3,'[',''''),']','''') " & _
                                        ''   " set @string3 = left(@string3, len(@string3) -1) " & _
                                        ''     " set @string4 = replace(@string3,'''','''''') " & _
                                        ''   " select @string3 [Ouput] , @string4 [quotes] "
                                        ''oRset.DoQuery(sSQL)
                                        ''sIN = oRset.Fields.Item("Ouput").Value
                                        ''sQuotes = oRset.Fields.Item("quotes").Value

                                        sIN = p_Dimensionrules
                                        sQuotes = Replace(p_Dimensionrules, "'", "''")
                                        oGridT2 = oForm.Items.Item("Item_23").Specific
                                        sNow = CStr(Now.Year) & sSplitM(0).PadLeft(2, "0"c) & "01"


                                        sSQL = "DECLARE @cols AS NVARCHAR(MAX),    @query  AS VARCHAR(max) , @query1  AS VARCHAR(max),  @cols1 as nvarchar(max) " & _
                                              "select @cols = STUFF((SELECT distinct  ',' + QUOTENAME(cast(T1.PrcCode as nvarchar(100))) " & _
                   " from OOCR T0  INNER JOIN OCR1 T1 ON T0.[OcrCode] = T1.[OcrCode]  where T0.OcrCode in (" & sIN & ") and T1.[ValidFrom] <= '" & sNow & "' and  (T1.[ValidTo] >= '" & sNow & "' or  isnull(T1.[ValidTo],'') = '')" & _
                   " and T1.PrcCode in (" & p_Summaryreport & ") " & _
           " FOR XML PATH(''), TYPE " & _
           " ).value('.', 'NVARCHAR(MAX)') " & _
       "  ,1,1,'') " & _
    "select @cols1 = STUFF((SELECT '+ isnull(' + QUOTENAME(isnull(T1.PrcCode,0)) + ',0)' " & _
                   " from OOCR T0  INNER JOIN OCR1 T1 ON T0.[OcrCode] = T1.[OcrCode]  where T0.OcrCode in (" & sIN & ") and T1.[ValidFrom] <= '" & sNow & "' and  (T1.[ValidTo] >= '" & sNow & "' or  isnull(T1.[ValidTo],'') = '')" & _
                    " and T1.PrcCode in (" & p_Summaryreport & ") " & _
           " FOR XML PATH(''), TYPE " & _
           " ).value('.', 'NVARCHAR(MAX)') " & _
       "  ,1,1,'') " & _
   " set @query = cast('SELECT SeriesName , ''PU '' + cast(DOCNUM as varchar(30)) ''SAP Reference Number'' , PDocNum ''Payment Doc Num'', PDocDate ''Payment Doc Date'' ,  NUMATCARD ''Vendor Invoice Number'',TAXDATE ''Vendor Invoice Date'',DOCDATE ''SAP Posting Date'',  " & _
    " CARDNAME ''Vendor / Payee Name'',JRNLMEMO ''Journal Remark'',  [OcrCode3] ''Distribution Code'', SLPNAME ''Purchasing Department'', Creator [Document Creator] , " & _
    "plan_id [Expenses], Line [Total Bill],' + @cols + '  " & _
                                          " from (" & _
               " SELECT i.SeriesName, f.DOCNUM   , f.NUMATCARD  ,f.TAXDATE  ,f.DOCDATE  ,f.CARDNAME  ,f.JRNLMEMO  ,  b.[OcrCode3]  ,h.SLPNAME  ," & _
               "a.AcctCode + '' - '' + a.AcctName [plan_id] , e.[PrcCode] [U_AB_NONPROJECT], l.DocNum [PDocNum] , l.DocDate [PDocDate] , g.LineNum , " & _
                "  case when g.BaseType = 18 then -sum(isnull(g.LineTotal,0)) else sum(isnull(g.LineTotal,0)) end [Line] , " & _
                " case when g.BaseType = 18 then -sum( round(isnull(g.LineTotal,0) /(d.[OcrTotal] /e.[PrcAmount] ),3)) else sum( round(isnull(g.LineTotal,0) /(d.[OcrTotal] /e.[PrcAmount] ),3)) end AS total , j.[U_NAME] +'', '' + j.[USER_CODE][Creator] " & _
                " FROM JDT1 b JOIN OACT a ON a.AcctCode = b.Account " & _
    " JOIN OJDT c on b.[TransId] = c.[TransId] " & _
    " left join OOCR d on d.OcrCode = b.OcrCode3 " & _
    " join OCR1 e on d.[OcrCode] = e.[OcrCode] " & _
    " JOIN OPCH f on c.[TransId] = f.[TransId]  JOIN PCH1 g on g.[DocEntry] = f.[DocEntry] and g.AcctCode = b.Account  and  g.[OcrCode3] =  b.[OcrCode3] " & _
    "  LEFT JOIN OSLP h on f.SLPCODE=h.SLPCODE " & _
    "  LEFT JOIN NNM1 i on i.Series=f.Series " & _
    "   LEFT JOIN OUSR j on  f.[usersign] = j.[USERID] " & _
    " left join   [VPM2] k on f.DocEntry = k.DocEntry left JOIN OVPM l ON k.[DocNum] = l.[DocEntry] " & _
    " where  b.OcrCode3 in (" & sQuotes & ") and a.GroupMask = 5  and isnull(c.Indicator,'''') <> ''CA'' and " & _
    "	b.Account >= ''" & sSplitG(0) & "'' and b.Account <= ''" & sSplitG(1) & "'' and c.transType IN (''18'') and month(c.refdate) >= " & sSplitM(0) & " and month(c.refdate) <= " & sSplitM(1) & " and year(c.refdate) = " & Now.Year & " " & _
    " and  e.[ValidFrom] <= ''" & sNow & "'' and (e.[ValidTo] >= ''" & sNow & "'' or  isnull(e.[ValidTo],'''') = '''')" & _
     " and e.PrcCode not in (SELECT T0.[U_PrcCode] FROM [dbo].[@AB_COSTALLOCATION]  T0 WHERE T0.[U_Transid] = c.TransId and T0.U_OcrCode3 = b.OcrCode3)  and (case when isnull(l.DocNum,'''') = '''' then ''N'' else l.Canceled end)=''N'' " & _
             "  group by a.AcctCode,a.AcctName,e.[PrcCode] , f.DOCNUM   , f.NUMATCARD  ,f.TAXDATE  ,f.DOCDATE  ,f.CARDNAME  ,f.JRNLMEMO  ,  b.[OcrCode3]  ,h.SLPNAME , g.LineNum, i.SeriesName ,j.[U_NAME] , j.[USER_CODE], l.DocNum , l.DocDate , g.BaseType  ) x pivot ( sum(total) " & _
             "   for U_AB_NONPROJECT in (' + @cols + ') " & _
           " ) p  " & _
               " union all SELECT SeriesName, ''PC '' + cast(DOCNUM as varchar(30)) ''SAP Reference Number'' , '''' ''Payment Doc Num'', NULL ''Payment Doc Date'',  NUMATCARD ''Vendor Invoice Number'',TAXDATE ''Vendor Invoice Date'',DOCDATE ''SAP Posting Date'',  " & _
    " CARDNAME ''Vendor / Payee Name'',JRNLMEMO ''Journal Remark'',  [OcrCode3] ''Distribution Code'', SLPNAME ''Purchasing Department'', Creator [Document Creator], " & _
    "plan_id [Expenses], Line [Total Bill],' + @cols + '  " & _
                                          " from (" & _
               " SELECT i.SeriesName, f.DOCNUM   , f.NUMATCARD  ,f.TAXDATE  ,f.DOCDATE  ,f.CARDNAME  ,f.JRNLMEMO  ,  b.[OcrCode3]  ,h.SLPNAME  , j.[U_NAME] +'', '' + j.[USER_CODE][Creator] ," & _
               "a.AcctCode + '' - '' + a.AcctName [plan_id] , e.[PrcCode] [U_AB_NONPROJECT], g.LineNum, " & _
                                        " case when g.BaseType = ''19'' then   sum(ISNULL( g.LineTotal,0))  else  - sum(ISNULL( g.LineTotal,0))  end [Line] , " & _
                " case when g.BaseType = ''19'' then  sum( round((isnull( g.LineTotal,0)) /(d.[OcrTotal] /e.[PrcAmount] ),3)) else - sum( round((isnull( g.LineTotal,0)) /(d.[OcrTotal] /e.[PrcAmount] ),3)) end AS total " & _
                " FROM JDT1 b JOIN OACT a ON a.AcctCode = b.Account " & _
    " JOIN OJDT c on b.[TransId] = c.[TransId] " & _
    " left join OOCR d on d.OcrCode = b.OcrCode3 " & _
    " join OCR1 e on d.[OcrCode] = e.[OcrCode] " & _
    " JOIN ORPC f on c.[TransId] = f.[TransId]  JOIN RPC1 g on g.[DocEntry] = f.[DocEntry] and g.AcctCode = b.Account  and  g.[OcrCode3] =  b.[OcrCode3] " & _
    "  LEFT JOIN OSLP h on f.SLPCODE=h.SLPCODE " & _
    "  LEFT JOIN NNM1 i on i.Series=f.Series " & _
    "   LEFT JOIN OUSR j on  f.[usersign] = j.[USERID] " & _
    " where  b.OcrCode3 in (" & sQuotes & ") and a.GroupMask = 5  and isnull(c.Indicator,'''') <> ''CA'' and " & _
    "	b.Account >= ''" & sSplitG(0) & "'' and b.Account <= ''" & sSplitG(1) & "'' and c.transType IN (''19'') and month(c.refdate) >= " & sSplitM(0) & " and month(c.refdate) <= " & sSplitM(1) & " and year(c.refdate) = " & Now.Year & " " & _
    " and  e.[ValidFrom] <= ''" & sNow & "'' and (e.[ValidTo] >= ''" & sNow & "'' or  isnull(e.[ValidTo],'''') = '''')" & _
     " and e.PrcCode not in (SELECT T0.[U_PrcCode] FROM [dbo].[@AB_COSTALLOCATION]  T0 WHERE T0.[U_Transid] = c.TransId and T0.U_OcrCode3 = b.OcrCode3) " & _
             "  group by a.AcctCode,a.AcctName,e.[PrcCode] , f.DOCNUM   , f.NUMATCARD  ,f.TAXDATE  ,f.DOCDATE  ,f.CARDNAME  ,f.JRNLMEMO  ,  b.[OcrCode3]  ,h.SLPNAME , g.LineNum , i.SeriesName ,j.[U_NAME] , j.[USER_CODE] ,g.BaseType ) x pivot ( sum(total) " & _
             "   for U_AB_NONPROJECT in (' + @cols + ') " & _
           " ) p ' as nvarchar(max)) " & _
  " set @query1 = cast( ' union all SELECT SeriesName, ''JE '' + cast(Number as varchar(30)) ''SAP Reference Number''  , '''' ''Payment Doc Num'', NULL ''Payment Doc Date'' ,REF2 ''Vendor Invoice Number'' , TAXDATE ''Vendor Invoice Date'' ,REFDATE ''SAP Posting Date'' " & _
        " ,''JE'' ''Vendor / Payee Name'',MEMO ''Journal Remark'' ,  [OcrCode3] ''Distribution Code'', '''' ''Purchasing Department'', Creator [Document Creator], plan_id [Expenses],  Line [Total Bill],' + @cols + '  " & _
        " from (" & _
               " SELECT i.SeriesName, c.Number   , b.REF2  ,b.TAXDATE  ,c.REFDATE  ,c.MEMO  ,  b.[OcrCode3]  , a.AcctCode + '' - '' + a.AcctName [plan_id] , e.[PrcCode] [U_AB_NONPROJECT] , j.[U_NAME] +'', '' + j.[USER_CODE][Creator] ," & _
                "   sum( round((ISNULL(b.Debit,0) - ISNULL(b.Credit,0)) /(d.[OcrTotal] /e.[PrcAmount] ),3)) AS total , sum(ISNULL(b.Debit,0) - ISNULL(b.Credit,0)) [Line] " & _
                " FROM JDT1 b JOIN OACT a ON a.AcctCode = b.Account " & _
    " JOIN OJDT c on b.[TransId] = c.[TransId] " & _
    " left join OOCR d on d.OcrCode = b.OcrCode3 " & _
    " join OCR1 e on d.[OcrCode] = e.[OcrCode] " & _
    "  LEFT JOIN NNM1 i on i.Series=c.Series " & _
    "  LEFT JOIN OUSR j on  c.[usersign] = j.[USERID]" & _
    " where  b.OcrCode3 in (" & sQuotes & ") and a.GroupMask = 5  and isnull(c.Indicator,'''') <> ''CA'' and " & _
    "	b.Account >= ''" & sSplitG(0) & "'' and b.Account <= ''" & sSplitG(1) & "'' and c.transType IN (''30'',''1470000071'') and month(c.refdate) >= " & sSplitM(0) & " and month(c.refdate) <= " & sSplitM(1) & " and year(c.refdate) = " & Now.Year & " " & _
    " and  e.[ValidFrom] <= ''" & sNow & "'' and (e.[ValidTo] >= ''" & sNow & "'' or  isnull(e.[ValidTo],'''') = '''')" & _
     " and e.PrcCode not in (SELECT T0.[U_PrcCode] FROM [dbo].[@AB_COSTALLOCATION]  T0 WHERE T0.[U_Transid] = c.TransId and T0.U_OcrCode3 = b.OcrCode3) " & _
    " group by c.Number   , b.REF2  ,b.TAXDATE  ,c.REFDATE  ,c.MEMO  ,  b.[OcrCode3]  , a.AcctCode , a.AcctName , e.[PrcCode],i.SeriesName,j.[U_NAME] , j.[USER_CODE] " & _
             "   ) x pivot ( sum(total) " & _
             "   for U_AB_NONPROJECT in (' + @cols + ') " & _
           " ) p " & _
           " UNION ALL SELECT  SeriesName, ''PS '' + cast(DOCNUM as varchar(30))''SAP Reference Number'' , '''' ''Payment Doc Num'', NULL ''Payment Doc Date'' ,'''' ''Vendor Invoice Number'',TAXDATE ''Vendor Invoice Date'',DOCDATE ''SAP Posting Date'',  " & _
    " Address ''Vendor / Payee Name'',JRNLMEMO ''Journal Remark'',  [OcrCode3] ''Distribution Code'', '''' ''Purchasing Department'', Creator [Document Creator], " & _
    "plan_id [Expenses], Line [Total Bill],' + @cols + '  " & _
                                          " from (" & _
               " SELECT i.SeriesName,f.DOCNUM   ,f.TAXDATE  ,f.DOCDATE   ,f.JRNLMEMO  ,  b.[OcrCode3]   ," & _
               "a.AcctCode + '' - '' + a.AcctName [plan_id] , e.[PrcCode] [U_AB_NONPROJECT],  f.[Address] , g.LineId,  j.[U_NAME] +'', '' + j.[USER_CODE][Creator] ," & _
                "  sum(isnull(g.SumApplied,0)) [Line] , sum( round(isnull(g.SumApplied,0) /(d.[OcrTotal] /e.[PrcAmount] ),3)) AS total " & _
                " FROM JDT1 b JOIN OACT a ON a.AcctCode = b.Account " & _
    " JOIN OJDT c on b.[TransId] = c.[TransId] " & _
    " left join OOCR d on d.OcrCode = b.OcrCode3 " & _
    " join OCR1 e on d.[OcrCode] = e.[OcrCode] " & _
    " JOIN OVPM f on c.[TransId] = f.[TransId]  JOIN VPM4 g on g.[DocNum] = f.[DocNum] and g.AcctCode = b.Account  and  g.[OcrCode3] =  b.[OcrCode3] " & _
     "  LEFT JOIN NNM1 i on i.Series=f.Series " & _
      "   LEFT JOIN OUSR j on  f.[usersign] = j.[USERID] " & _
     " where  b.OcrCode3 in (" & sQuotes & ") and a.GroupMask = 5  and isnull(c.Indicator,'''') <> ''CA'' and " & _
    "	b.Account >= ''" & sSplitG(0) & "'' and b.Account <= ''" & sSplitG(1) & "'' and c.transType IN (''46'') and month(c.refdate) >= " & sSplitM(0) & " and month(c.refdate) <= " & sSplitM(1) & " and year(c.refdate) = " & Now.Year & " " & _
    " and  e.[ValidFrom] <= ''" & sNow & "'' and (e.[ValidTo] >= ''" & sNow & "'' or  isnull(e.[ValidTo],'''') = '''')" & _
     " and e.PrcCode not in (SELECT T0.[U_PrcCode] FROM [dbo].[@AB_COSTALLOCATION]  T0 WHERE T0.[U_Transid] = c.TransId and T0.U_OcrCode3 = b.OcrCode3) " & _
     "  and f.DocType = ''A'' " & _
             "  group by a.AcctCode,a.AcctName,e.[PrcCode] , f.DOCNUM ,f.TAXDATE  ,f.DOCDATE  ,f.JRNLMEMO  ,  b.[OcrCode3]  , g.LineId , f.[Address],i.SeriesName ,j.[U_NAME] , j.[USER_CODE] ) x pivot ( sum(total) " & _
             "   for U_AB_NONPROJECT in (' + @cols + ') " & _
           " ) p " & _
           " UNION ALL SELECT  SeriesName, ''PS '' + cast(DOCNUM as varchar(30))''SAP Reference Number'' , '''' ''Payment Doc Num'', NULL ''Payment Doc Date'' ,'''' ''Vendor Invoice Number'',TAXDATE ''Vendor Invoice Date'',DOCDATE ''SAP Posting Date'',  " & _
    " Address ''Vendor / Payee Name'',JRNLMEMO ''Journal Remark'',  [OcrCode3] ''Distribution Code'', '''' ''Purchasing Department'', Creator [Document Creator], " & _
    "plan_id [Expenses], Line [Total Bill],' + @cols + '  " & _
                                          " from (" & _
               " SELECT i.SeriesName,f.DOCNUM   ,f.CancelDate [TAXDATE]  ,f.CancelDate [DOCDATE]   ,f.JRNLMEMO  ,  b.[OcrCode3]   ," & _
               "a.AcctCode + '' - '' + a.AcctName [plan_id] , e.[PrcCode] [U_AB_NONPROJECT],  f.[Address] , g.LineId ,  j.[U_NAME] +'', '' + j.[USER_CODE][Creator] ," & _
                "  -sum(isnull(g.SumApplied,0)) [Line] , -sum( round(isnull(g.SumApplied,0) /(d.[OcrTotal] /e.[PrcAmount] ),3)) AS total " & _
                " FROM JDT1 b JOIN OACT a ON a.AcctCode = b.Account " & _
    " JOIN OJDT c on b.[TransId] = c.[TransId] " & _
    " left join OOCR d on d.OcrCode = b.OcrCode3 " & _
    " join OCR1 e on d.[OcrCode] = e.[OcrCode] " & _
    " JOIN OVPM f on c.[TransId] = f.[TransId]  JOIN VPM4 g on g.[DocNum] = f.[DocNum] and g.AcctCode = b.Account   and  g.[OcrCode3] =  b.[OcrCode3] " & _
     "  LEFT JOIN NNM1 i on i.Series=f.Series " & _
      "   LEFT JOIN OUSR j on  f.[usersign] = j.[USERID] " & _
     " where  b.OcrCode3 in (" & sQuotes & ") and a.GroupMask = 5  and isnull(c.Indicator,'''') <> ''CA'' and " & _
    "	b.Account >= ''" & sSplitG(0) & "'' and b.Account <= ''" & sSplitG(1) & "'' and c.transType IN (''46'') and month(f.CancelDate) >= " & sSplitM(0) & " and month(f.CancelDate) <= " & sSplitM(1) & " and year(f.CancelDate) = " & Now.Year & " " & _
    " and  e.[ValidFrom] <= ''" & sNow & "'' and (e.[ValidTo] >= ''" & sNow & "'' or  isnull(e.[ValidTo],'''') = '''')" & _
     " and e.PrcCode not in (SELECT T0.[U_PrcCode] FROM [dbo].[@AB_COSTALLOCATION]  T0 WHERE T0.[U_Transid] = c.TransId and T0.U_OcrCode3 = b.OcrCode3) " & _
     "  and f.DocType = ''A'' and f.Canceled = ''Y''" & _
             "  group by a.AcctCode,a.AcctName,e.[PrcCode] , f.DOCNUM ,f.CancelDate ,f.JRNLMEMO  ,  b.[OcrCode3]  , g.LineId , f.[Address],i.SeriesName ,j.[U_NAME] , j.[USER_CODE]  ) x pivot ( sum(total) " & _
             "   for U_AB_NONPROJECT in (' + @cols + ') " & _
           " ) p ' as nvarchar(max)) " & _
  " execute (@query + @query1)"


                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("SummaryReport show button click " & sSQL, sFuncName)
                                        Try
                                            oForm.DataSources.DataTables.Add("MyDataTable1")
                                        Catch ex As Exception
                                        End Try
                                        oGridT2.DataTable = oForm.DataSources.DataTables.Item("MyDataTable1")
                                        oForm.DataSources.DataTables.Item(1).ExecuteQuery(sSQL)
                                        oGridT2.CollapseLevel = 2
                                        oGridT2.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
                                        Dim brow As Integer = oGridT2.DataTable.Rows.Count


                                        sSQL = "select * from (SELECT a.AcctCode [Account], a.AcctName [ShortName], e.U_AB_NONPROJECT [U_AB_NONPROJECT],  isnull(e.Project,'') [Project] ,e1.PrcCode [OcrCode3] , f.PrcName [U_AB_OUName]," & _
                                        " round((ISNULL(b.Debit,0) - ISNULL(b.Credit,0)) /(d1.[OcrTotal] /e1.[PrcAmount] ),3) AS Credit, isnull(b.LineMemo,'') [U_AB_REMARKS] , '' U_AB_PARTNER, b.TransId, d.DocEntry [BaseRef] , 'AP' 'LineMemo' " & _
                                        " , isnull(e.OcrCode2,'') [BU], isnull(e.OcrCode,'') [LOS] , isnull( b.VatGroup  ,'') [VatGroup] , b.OcrCode3 [ContraAct]  " & _
                                          "FROM JDT1 b JOIN OACT a ON a.AcctCode = b.Account  " & _
                                          " JOIN OJDT c on b.[TransId] = c.[TransId]  JOIN OPCH d on c.[TransId] = d.[TransId]  " & _
                                          " JOIN PCH1 e on e.[DocEntry] = d.[DocEntry] and e.AcctCode = b.Account  and e.[OcrCode3] = b.[OcrCode3] " & _
                                           " left join OOCR d1 on d1.OcrCode = b.OcrCode3 " & _
                                          " join OCR1 e1 on d1.[OcrCode] = e1.[OcrCode] left outer join OPRC f on f.PrcCode = e1.PrcCode " & _
                                          " where  b.OcrCode3 in (" & sIN & ") and a.GroupMask = 5  and isnull(d.Indicator,'') <> 'CA' and " & _
                                         "	b.Account >= '" & sSplitG(0) & "' and b.Account <= '" & sSplitG(1) & "' and c.transType IN ('18') and month(c.refdate) >= " & sSplitM(0) & " and month(c.refdate) <= " & sSplitM(1) & " and year(c.refdate) = " & Now.Year & " " & _
                                          " and  e1.[ValidFrom] <= '" & sNow & "' and (e1.[ValidTo] >= '" & sNow & "' or  isnull(e1.[ValidTo],'') = '')" & _
                                           " and e1.PrcCode in (" & p_Summaryreport & ")" & _
                                            " and e1.PrcCode not in (SELECT T0.[U_PrcCode] FROM [dbo].[@AB_COSTALLOCATION]  T0 WHERE T0.[U_Transid] = c.TransId and T0.U_OcrCode3 = b.OcrCode3) " & _
                                            " group by " & _
                                           "  a.AcctCode  , a.AcctName  , e.U_AB_NONPROJECT  ,  isnull(e.Project,'') ,e1.PrcCode  , f.PrcName  ,  b.Debit, b.Credit,d1.[OcrTotal],e1.[PrcAmount], isnull(b.LineMemo,'') ,  b.TransId, d.DocEntry , isnull(e.OcrCode2,'')  , isnull(e.OcrCode ,'') ,isnull( b.VatGroup  ,'') , b.OcrCode3 " & _
                                                                               "  UNION ALL " & _
                                          " SELECT   a.AcctCode [Account], a.AcctName [ShortName], e.U_AB_NONPROJECT [U_AB_NONPROJECT],  isnull(e.Project,'') [Project], e1.PrcCode [OcrCode3] , f.PrcName [U_AB_OUName] , " & _
                                          " round((ISNULL(b.Debit,0) - ISNULL(b.Credit,0)) /(d1.[OcrTotal] /e1.[PrcAmount] ),3) AS Credit,isnull( b.LineMemo,'') [U_AB_REMARKS] , '' U_AB_PARTNER ,  b.TransId, d.DocEntry [BaseRef] , 'CN' 'LineMemo'  " & _
                                          " , isnull(e.OcrCode2,'') [BU], isnull(e.OcrCode,'') [LOS] , isnull( b.VatGroup  ,'') [VatGroup] , b.OcrCode3 [ContraAct] " & _
                                          "  FROM JDT1 b JOIN OACT a  ON a.AcctCode = b.Account  JOIN OJDT c on b.[TransId] = c.[TransId]  " & _
                                          " JOIN ORPC d on c.[TransId] = d.[TransId]  JOIN RPC1 e on e.[DocEntry] = d.[DocEntry]  and e.AcctCode = b.Account  and e.[OcrCode3] = b.[OcrCode3] " & _
                                           " left join OOCR d1 on d1.OcrCode = b.OcrCode3 " & _
                                          " join OCR1 e1 on d1.[OcrCode] = e1.[OcrCode]  left outer join OPRC f on f.PrcCode = e1.PrcCode" & _
                                          " where  b.OcrCode3 in (" & sIN & ") and a.GroupMask = 5  and isnull(d.Indicator,'') <> 'CA' and " & _
                                         "	b.Account >= '" & sSplitG(0) & "' and b.Account <= '" & sSplitG(1) & "' and c.transType IN ('19') and month(c.refdate) >= " & sSplitM(0) & " and month(c.refdate) <= " & sSplitM(1) & " and year(c.refdate) = " & Now.Year & " " & _
                                          " and  e1.[ValidFrom] <= '" & sNow & "' and (e1.[ValidTo] >= '" & sNow & "' or  isnull(e1.[ValidTo],'') = '')" & _
                                           " and e1.PrcCode in (" & p_Summaryreport & ")" & _
                                         " and e1.PrcCode not in (SELECT T0.[U_PrcCode] FROM [dbo].[@AB_COSTALLOCATION]  T0 WHERE T0.[U_Transid] = c.TransId and T0.U_OcrCode3 = b.OcrCode3) " & _
                                         " group by " & _
                                          "  a.AcctCode  , a.AcctName  , e.U_AB_NONPROJECT  , isnull( e.Project,'') ,e1.PrcCode  , f.PrcName  ,  b.Debit, b.Credit,d1.[OcrTotal],e1.[PrcAmount],isnull( b.LineMemo,'') ,  b.TransId, d.DocEntry  , isnull(e.OcrCode2,'')  , isnull(e.OcrCode ,'') ,isnull( b.VatGroup  ,'') , b.OcrCode3 " & _
                                                                                  "  UNION ALL " & _
                                         " SELECT   a.AcctCode [Account], a.AcctName [ShortName], b.U_AB_NONPROJECT [U_AB_NONPROJECT],  isnull(b.Project,'') [Project] ,e1.PrcCode [OcrCode3] , f.PrcName [U_AB_OUName] , " & _
                                         " round((ISNULL(b.Debit,0) - ISNULL(b.Credit,0)) /(d1.[OcrTotal] /e1.[PrcAmount] ),3) AS Credit, isnull(b.LineMemo,'') [U_AB_REMARKS] ,  isnull(g.Name,'')  U_AB_PARTNER ,  b.TransId, 0 [BaseRef] , case when c.transType = '30' then case when left(a.AcctCode,4) = '7212' or left(a.AcctCode,4) = '7213' then  'DEP' else 'JE' end  else 'DEP' end 'LineMemo', isnull(b.OcrCode2,'') [BU] , isnull(b.ProfitCode,'') [LOS] , isnull( b.VatGroup  ,'') [VatGroup] , b.OcrCode3 [ContraAct]  " & _
                                          "  FROM JDT1 b JOIN OACT a  ON a.AcctCode = b.Account  JOIN OJDT c on b.[TransId] = c.[TransId]  " & _
                                          " left join OOCR d1 on d1.OcrCode = b.OcrCode3 " & _
                                          " join OCR1 e1 on d1.[OcrCode] = e1.[OcrCode]  left outer join OPRC f on f.PrcCode = e1.PrcCode" & _
                                          " left outer join [@AB_PARTNER] g on g.Code = b.U_AB_PARTNER" & _
                                          " where  b.OcrCode3 in (" & sIN & ") and a.GroupMask = 5  and isnull(c.Indicator,'') <> 'CA' and " & _
                                         "	b.Account >= '" & sSplitG(0) & "' and b.Account <= '" & sSplitG(1) & "' and c.transType IN ('30','1470000071') and month(c.refdate) >= " & sSplitM(0) & " and month(c.refdate) <= " & sSplitM(1) & " and year(c.refdate) = " & Now.Year & " " & _
                                          " and  e1.[ValidFrom] <= '" & sNow & "' and (e1.[ValidTo] >= '" & sNow & "' or  isnull(e1.[ValidTo],'') = '')" & _
                                           " and e1.PrcCode in (" & p_Summaryreport & ") " & _
 " and e1.PrcCode not in (SELECT T0.[U_PrcCode] FROM [dbo].[@AB_COSTALLOCATION]  T0 WHERE T0.[U_Transid] = c.TransId and T0.U_OcrCode3 = b.OcrCode3) " & _
  " group by " & _
                                           "  a.AcctCode  , a.AcctName  , b.U_AB_NONPROJECT  , isnull( b.Project,'') ,e1.PrcCode  , f.PrcName  ,  b.Debit, b.Credit,d1.[OcrTotal],e1.[PrcAmount],isnull( b.LineMemo,'') ,  b.TransId, isnull( b.OcrCode2,'')   , isnull(b.ProfitCode,''), isnull(g.Name,''),isnull( b.VatGroup  ,'') ,b.OcrCode3 ,c.transType " & _
 " UNION ALL SELECT a.AcctCode [Account], a.AcctName [ShortName], e.U_AB_NONPROJECT [U_AB_NONPROJECT],  isnull(e.Project,'') [Project] ,e1.PrcCode [OcrCode3] , f.PrcName [U_AB_OUName]," & _
    " round((ISNULL(b.Debit,0) - ISNULL(b.Credit,0)) /(d1.[OcrTotal] /e1.[PrcAmount] ),3) AS Credit, isnull(b.LineMemo,'') [U_AB_REMARKS] , '' U_AB_PARTNER, b.TransId, d.DocEntry [BaseRef] , 'AP' 'LineMemo' " & _
                                        " , isnull(e.OcrCode2,'') [BU], isnull(e.OcrCode,'') [LOS] , isnull( b.VatGroup  ,'') [VatGroup] , b.OcrCode3 [ContraAct]  " & _
    "FROM JDT1 b JOIN OACT a ON a.AcctCode = b.Account  " & _
    " JOIN OJDT c on b.[TransId] = c.[TransId]  JOIN OVPM d on c.[TransId] = d.[TransId]  " & _
    " JOIN VPM4 e on e.[DocNum] = d.[DocNum] and e.AcctCode = b.Account  and e.[OcrCode3] = b.[OcrCode3] " & _
     " left join OOCR d1 on d1.OcrCode = b.OcrCode3 " & _
    " join OCR1 e1 on d1.[OcrCode] = e1.[OcrCode] left outer join OPRC f on f.PrcCode = e1.PrcCode " & _
    " where  b.OcrCode3 in (" & sIN & ") and a.GroupMask = 5  and d.DocType = 'A' and " & _
   "	b.Account >= '" & sSplitG(0) & "' and b.Account <= '" & sSplitG(1) & "' and c.transType IN ('46') and month(c.refdate) >= " & sSplitM(0) & " and month(c.refdate) <= " & sSplitM(1) & " and year(c.refdate) = " & Now.Year & " " & _
    " and  e1.[ValidFrom] <= '" & sNow & "' and (e1.[ValidTo] >= '" & sNow & "' or  isnull(e1.[ValidTo],'') = '')" & _
    " and e1.PrcCode in (" & p_Summaryreport & ") " & _
    " and e1.PrcCode not in (SELECT T0.[U_PrcCode] FROM [dbo].[@AB_COSTALLOCATION]  T0 WHERE T0.[U_Transid] = c.TransId and T0.U_OcrCode3 = b.OcrCode3) " & _
    " group by " & _
    "  a.AcctCode  , a.AcctName  , e.U_AB_NONPROJECT  ,  isnull(e.Project,'') ,e1.PrcCode  , f.PrcName  ,  b.Debit, b.Credit,d1.[OcrTotal],e1.[PrcAmount], isnull(b.LineMemo,'') ,  b.TransId, d.DocEntry , isnull(e.OcrCode2,'')  , isnull(e.OcrCode ,'') ,isnull( b.VatGroup  ,'') ,b.OcrCode3 " & _
     " UNION ALL SELECT a.AcctCode [Account], a.AcctName [ShortName], e.U_AB_NONPROJECT [U_AB_NONPROJECT],  isnull(e.Project,'') [Project] ,e1.PrcCode [OcrCode3] , f.PrcName [U_AB_OUName]," & _
    " -round((ISNULL(b.Debit,0) - ISNULL(b.Credit,0)) /(d1.[OcrTotal] /e1.[PrcAmount] ),3) AS Credit, isnull(b.LineMemo,'') [U_AB_REMARKS] , '' U_AB_PARTNER, b.TransId, d.DocEntry [BaseRef] , 'AP' 'LineMemo' " & _
                                        " , isnull(e.OcrCode2,'') [BU], isnull(e.OcrCode,'') [LOS] , isnull( b.VatGroup  ,'') [VatGroup] , b.OcrCode3 [ContraAct]  " & _
    "FROM JDT1 b JOIN OACT a ON a.AcctCode = b.Account  " & _
    " JOIN OJDT c on b.[TransId] = c.[TransId]  JOIN OVPM d on c.[TransId] = d.[TransId]  " & _
    " JOIN VPM4 e on e.[DocNum] = d.[DocNum] and e.AcctCode = b.Account  and e.[OcrCode3] = b.[OcrCode3] " & _
     " left join OOCR d1 on d1.OcrCode = b.OcrCode3 " & _
    " join OCR1 e1 on d1.[OcrCode] = e1.[OcrCode] left outer join OPRC f on f.PrcCode = e1.PrcCode " & _
    " where  b.OcrCode3 in (" & sIN & ") and a.GroupMask = 5  and d.DocType = 'A' and d.Canceled = 'Y' and " & _
   "	b.Account >= '" & sSplitG(0) & "' and b.Account <= '" & sSplitG(1) & "' and c.transType IN ('46') and month(d.CancelDate) >= " & sSplitM(0) & " and month(d.CancelDate) <= " & sSplitM(1) & " and year(d.CancelDate) = " & Now.Year & " " & _
    " and  e1.[ValidFrom] <= '" & sNow & "' and (e1.[ValidTo] >= '" & sNow & "' or  isnull(e1.[ValidTo],'') = '')" & _
    " and e1.PrcCode in (" & p_Summaryreport & ") " & _
    " and e1.PrcCode not in (SELECT T0.[U_PrcCode] FROM [dbo].[@AB_COSTALLOCATION]  T0 WHERE T0.[U_Transid] = c.TransId and T0.U_OcrCode3 = b.OcrCode3) " & _
    " group by " & _
    "  a.AcctCode  , a.AcctName  , e.U_AB_NONPROJECT  ,  isnull(e.Project,'') ,e1.PrcCode  , f.PrcName  ,  b.Debit, b.Credit,d1.[OcrTotal],e1.[PrcAmount], isnull(b.LineMemo,'') ,  b.TransId, d.DocEntry , isnull(e.OcrCode2,'')  , isnull(e.OcrCode ,'') ,isnull( b.VatGroup  ,'') ,b.OcrCode3 " & _
    ")x order by x.OcrCode3 , cast(x.Account as integer)"



                                        Try
                                            oForm.DataSources.DataTables.Add("JDT1")
                                        Catch ex As Exception
                                        End Try
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Tab 3 - Show butto click " & sSQL, sFuncName)
                                        Dim oSApDT As SAPbouiCOM.DataTable

                                        oSApDT = oForm.DataSources.DataTables.Item("JDT1")
                                        '' oForm.DataSources.DataTables.Item("JDT1").ExecuteQuery(sSQL)
                                        oSApDT.ExecuteQuery(sSQL)
                                        oMatrix.CommonSetting.SeparateLine(oMatrix.RowCount, 245, SAPbouiCOM.BoSeparateLineType.slt_Bottom, False)
                                        oMatrix.CommonSetting.SeparateLine(oMatrix.RowCount - 1, 245, SAPbouiCOM.BoSeparateLineType.slt_Bottom, False)
                                        oMatrix.CommonSetting.SeparateLine(oMatrix.RowCount - 2, 245, SAPbouiCOM.BoSeparateLineType.slt_Bottom, False)
                                        '' oMatrix.CommonSetting.SetRowFontSize(oMatrix.RowCount, 9)
                                        oMatrix.Clear()

                                        oForm.Items.Item("Item_24").Specific.columns.item("Col_0").databind.bind("JDT1", "Account")
                                        oForm.Items.Item("Item_24").Specific.columns.item("Col_1").databind.bind("JDT1", "ShortName")
                                        oForm.Items.Item("Item_24").Specific.columns.item("Col_2").databind.bind("JDT1", "U_AB_NONPROJECT")
                                        oForm.Items.Item("Item_24").Specific.columns.item("Col_3").databind.bind("JDT1", "Project")
                                        oForm.Items.Item("Item_24").Specific.columns.item("Col_4").databind.bind("JDT1", "OcrCode3")
                                        oForm.Items.Item("Item_24").Specific.columns.item("Col_5").databind.bind("JDT1", "U_AB_OUName")
                                        ''oForm.Items.Item("Item_24").Specific.columns.item("Col_6").databind.bind("JDT1", "U_AB_PARTNER")
                                        oForm.Items.Item("Item_24").Specific.columns.item("Col_7").databind.bind("JDT1", "Credit")
                                        oForm.Items.Item("Item_24").Specific.columns.item("Col_8").databind.bind("JDT1", "U_AB_REMARKS")
                                        oForm.Items.Item("Item_24").Specific.columns.item("Col_9").databind.bind("JDT1", "TransId")
                                        oForm.Items.Item("Item_24").Specific.columns.item("Col_10").databind.bind("JDT1", "BaseRef")
                                        oForm.Items.Item("Item_24").Specific.columns.item("Col_11").databind.bind("JDT1", "LineMemo")
                                        oForm.Items.Item("Item_24").Specific.columns.item("V_1").databind.bind("JDT1", "BU")
                                        oForm.Items.Item("Item_24").Specific.columns.item("V_0").databind.bind("JDT1", "LOS")
                                        oForm.Items.Item("Item_24").Specific.columns.item("V_3").databind.bind("JDT1", "U_AB_PARTNER")
                                        oForm.Items.Item("Item_24").Specific.columns.item("V_2").databind.bind("JDT1", "VatGroup")
                                        oForm.Items.Item("Item_24").Specific.columns.item("V_4").databind.bind("JDT1", "ContraAct")
                                        oForm.Items.Item("Item_24").Specific.LoadFromDataSource()
                                        oForm.Items.Item("Item_24").Specific.AutoResizeColumns()
                                        oMatrix.Columns.Item("Col_6").Editable = True

                                        Dim dsum As Double
                                        'For imjs As Integer = 0 To oSApDT.Rows.Count - 1
                                        '    dsum += Decimal.Round(oSApDT.GetValue("Credit", imjs), 2)
                                        'Next

                                        For imjs As Integer = 1 To oMatrix.RowCount
                                            dsum += CDbl(oMatrix.Columns.Item("Col_7").Cells.Item(imjs).Specific.String())
                                        Next
                                        oForm.DataSources.DataTables.Item(2).Clear()

                                        oMatrix.AddRow(2)
                                        oMatrix.CommonSetting.SeparateLine(oMatrix.RowCount, 245, SAPbouiCOM.BoSeparateLineType.slt_Bottom, True)
                                        oMatrix.AddRow(1)
                                        oMatrix.Columns.Item("Col_7").Cells.Item(oMatrix.RowCount).Specific.String = dsum
                                        ''oMatrix.CommonSetting.SetRowFontSize(oMatrix.RowCount, 11)
                                        oMatrix.Columns.Item("Col_5").Cells.Item(oMatrix.RowCount).Specific.String = "Total"
                                        oMatrix.CommonSetting.SeparateLine(oMatrix.RowCount, 245, SAPbouiCOM.BoSeparateLineType.slt_Bottom, True)


                                        sSplitM = p_Summaryreport.Split(",")
                                        oGridT2.DataTable.Rows.Add(1)
                                        oGridT2.DataTable.SetValue(0, brow, "Total")
                                        Dim bcount As Decimal = 0.0
                                        For imjs As Integer = 13 To oGridT2.Columns.Count - 1
                                            bcount = 0
                                            For imjd As Integer = 0 To brow - 1
                                                'p_Summaryreport
                                                If Not String.IsNullOrEmpty(oGridT2.DataTable.GetValue(imjs, imjd)) Then
                                                    bcount += Decimal.Round(oGridT2.DataTable.GetValue(imjs, imjd), 2)
                                                End If
                                            Next
                                            If sSplitM.Length > 1 Then
                                                oGridT2.DataTable.SetValue(imjs, brow, CDbl(bcount))
                                            ElseIf sSplitM.Length = 1 Then
                                                If imjs = oGridT2.Columns.Count - 1 Then
                                                    oGridT2.DataTable.SetValue(imjs, brow, dsum)
                                                Else
                                                    oGridT2.DataTable.SetValue(imjs, brow, CDbl(bcount))
                                                End If
                                            End If
                                        Next
                                        oGridT2.CommonSetting.SetRowBackColor(oGridT2.Rows.Count, 234)
                                        oGridT2.AutoResizeColumns()

                                        oForm.Freeze(False)
                                        p_oSBOApplication.SetStatusBarMessage("Information Loaded on the Summary Report .......!", SAPbouiCOM.BoMessageTime.bmt_Short, False)

                                    Catch ex As Exception
                                        oForm.Freeze(False)
                                        p_oSBOApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                        BubbleEvent = False
                                        sErrDesc = ex.Message
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                                        WriteToLogFile(Err.Description, sFuncName)
                                        ShowErr(sErrDesc)
                                    End Try
                                End If

                                If pVal.ItemUID = "Item_20x" Then
                                    Dim oDs As DataSet = Nothing
                                    Dim oDV As DataView = Nothing
                                    Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                                    Try
                                        SBO_Application.SetStatusBarMessage("Printing In Process ...! ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                        oDs = MAtrixToDataTable_E(oForm, sErrDesc)
                                        PrintCalling(oDs)
                                        '' ExportToExcel(oDT, System.Windows.Forms.Application.StartupPath & "\JE.xls")
                                    Catch ex As Exception
                                        p_oSBOApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                        BubbleEvent = False
                                        Exit Sub
                                    End Try

                                End If

                                If pVal.ItemUID = "Item_20" Then
                                    Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                                    Dim oDT As DataTable = Nothing
                                    Dim oDV As DataView = Nothing
                                    Dim oDVL As DataView = Nothing
                                    Dim oDVJV As DataView = Nothing
                                    Dim oDT_JV As DataTable = Nothing
                                    Dim oDT_Entity As DataTable = Nothing
                                    Dim oDICompany() As SAPbobsCOM.Company = Nothing
                                    Dim orset As SAPbobsCOM.Recordset = Nothing
                                    Dim irow As Integer = 0
                                    Dim sSQL As String = String.Empty
                                    Dim ilstday As Integer = 0
                                    Dim sNow As String = String.Empty
                                    Dim sRef As String = String.Empty
                                    Dim i_dSourceTable As DataTable = Nothing

                                    Try
                                        SBO_Application.SetStatusBarMessage("Attempting the function Journal Entry Creation ...! ", SAPbouiCOM.BoMessageTime.bmt_Long, False)
                                        sErrDesc = String.Empty
                                        oDV = MAtrixToDataTable(oForm, sErrDesc)

                                        For Each oddr As DataRowView In oDV
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug1(oddr("GL_Code") & "  " & oddr("GL_NameT") & "  " & oddr("OU") & "  " & oddr("Amount") & "  " & oddr("TGL"), "MAtrixToDataTable")
                                        Next

                                        If sErrDesc.Length > 0 Then Throw New ArgumentException(sErrDesc)
                                        oDVL = oDV

                                        If sErrDesc.Length > 0 Then Throw New ArgumentException(sErrDesc)
                                        oDT_Entity = oDV.ToTable(True, "EntityCode")
                                        ReDim oDICompany(oDT_Entity.Rows.Count - 1)
                                        orset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        ilstday = Date.DaysInMonth(Now.Year, oForm.Items.Item("Item_13").Specific.selected.description)
                                        sNow = CStr(Now.Year) & oForm.Items.Item("Item_13").Specific.selected.description.PadLeft(2, "0"c) & ilstday
                                        oDT_JV = oDVL.ToTable(True, "EntityCode")
                                        SBO_Application.SetStatusBarMessage("Identifying the Traget Entity ...! ", SAPbouiCOM.BoMessageTime.bmt_Short, False)

                                        For Each odrjv As DataRow In oDT_JV.Rows
                                            '' For Each odr As DataRow In oDT_Entity.Rows
                                            p_oSBOApplication.RemoveWindowsMessage(SAPbouiCOM.BoWindowsMessageType.bo_WM_TIMER, True)
                                            oDICompany(irow) = New SAPbobsCOM.Company
                                            oDVL.RowFilter = "EntityCode='" & odrjv("EntityCode") & "'"
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectToTargetCompany()", sFuncName)
                                            SBO_Application.SetStatusBarMessage("Connecting to the Target Company " & oDVL.Item(0)("EntityCode").ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                            If ConnectToTargetCompany(oDICompany(irow), oDVL.Item(0)("EntityCode").ToString(), oDVL.Item(0)("UName").ToString(), oDVL.Item(0)("Pass").ToString(), sErrDesc) <> RTN_SUCCESS Then
                                                '' If DisplayStatus(oForm, "Error " & sErrDesc, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                                Throw New ArgumentException(sErrDesc)
                                            End If
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starts JV Group ", sFuncName)
                                            oDVL.RowFilter = "JVGroup='JV' and EntityCode='" & odrjv("EntityCode") & "'"
                                            If oDVL.Count > 0 Then
                                                If p_oDICompany.InTransaction = False Then
                                                    p_oDICompany.StartTransaction()
                                                End If
                                                If oDICompany(irow).InTransaction = False Then
                                                    oDICompany(irow).StartTransaction()
                                                End If
                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling JournalEntry_Posting_JV_Source()", sFuncName)
                                                If JournalEntry_Posting_JV_Source(oDVL, p_oDICompany, sNow, sRef, sErrDesc) <> RTN_SUCCESS Then
                                                    If p_oDICompany.InTransaction Then
                                                        p_oDICompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                                    End If
                                                    For lCounter As Integer = 0 To UBound(oDICompany)
                                                        If Not oDICompany(lCounter) Is Nothing Then
                                                            If oDICompany(lCounter).Connected = True Then
                                                                If oDICompany(lCounter).InTransaction = True Then
                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback transaction on company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                    oDICompany(lCounter).EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                                                End If
                                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                oDICompany(lCounter).Disconnect()
                                                                oDICompany(lCounter) = Nothing
                                                            End If
                                                        End If
                                                    Next
                                                    Throw New ArgumentException(sErrDesc)
                                                End If
                                                '''''  Commenting on 22 SEP 2016
                                                ''i_dSourceTable = New DataTable
                                                '' ''  Dim oDV_Tmp As DataView = oDVL
                                                ''i_dSourceTable = oDVL.ToTable
                                                '' '' Dim dtGroup As DataTable = oDVL.ToTable(True, New String() {"JVGroup", "NewOU", "GL_NameT", "JV", "Base"})
                                                '' ''   Dim dtGroup As DataTable = oDVL.ToTable(True, New String() {i_sGroupByColumn, ii_sGroupByColumn, iii_sGroupByColumn, iiii_sGroupByColumn})
                                                ''Dim dtGroup As DataTable = oDVL.ToTable(True, "GL_NameT", "OU_BU_Budget", "Project", "NewOU", "Remarks", _
                                                ''                                         "TGL", "BU", "LOS", "Cat")
                                                ' ''adding column for the row count
                                                ''dtGroup.Columns.Add("Amount", GetType(Decimal))
                                                ' ''looping thru distinct values for the group, counting
                                                '' ''i_sGroupByColumn & " = '" & dr(i_sGroupByColumn) & "' AND " & ii_sGroupByColumn & " = '" & dr(ii_sGroupByColumn) & "' AND " & iii_sGroupByColumn & " = '" & dr(iii_sGroupByColumn) & "' AND " & iiii_sGroupByColumn & " = '" & dr(iiii_sGroupByColumn) & "'
                                                ''For Each dr As DataRow In dtGroup.Rows
                                                ''    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(dr("NewOU") & "   " & dr("GL_NameT"), sFuncName)
                                                ''    dr("Amount") = i_dSourceTable.Compute("Sum(Amount)", " GL_NameT = '" & dr("GL_NameT") & "' AND OU_BU_Budget = '" & dr("OU_BU_Budget") & "' AND Project = '" & dr("Project") & "' " & _
                                                ''                                       " AND NewOU = '" & dr("NewOU") & "' AND Remarks = '" & dr("Remarks").ToString & "'  " & _
                                                ''                                        "  AND TGL = '" & dr("TGL") & "' AND " & _
                                                ''                                        "  BU = '" & dr("BU") & "' AND LOS = '" & dr("LOS") & "' AND Cat = '" & dr("Cat") & "'")
                                                ''Next
                                                ''oDVJV = New DataView(dtGroup)
                                                ''If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling JournalEntry_Posting_JV_Target()", sFuncName)
                                                ''If JournalEntry_Posting_JV_Target(oDVJV, oDICompany(irow), sNow, sRef, sErrDesc) <> RTN_SUCCESS Then
                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling JournalEntry_Posting_JV_Target()", sFuncName)
                                                If JournalEntry_Posting_JV_Target(oDVL, oDICompany(irow), sNow, sRef, sErrDesc) <> RTN_SUCCESS Then

                                                    If p_oDICompany.InTransaction Then
                                                        p_oDICompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                                    End If
                                                    For lCounter As Integer = 0 To UBound(oDICompany)
                                                        If Not oDICompany(lCounter) Is Nothing Then
                                                            If oDICompany(lCounter).Connected = True Then
                                                                If oDICompany(lCounter).InTransaction = True Then
                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback transaction on company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                    oDICompany(lCounter).EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                                                End If
                                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                oDICompany(lCounter).Disconnect()
                                                                oDICompany(lCounter) = Nothing
                                                            End If
                                                        End If
                                                    Next
                                                    Throw New ArgumentException("In Target :- " & sErrDesc)
                                                End If
                                                sSQL = ""

                                                Dim oDT_CA As DataTable = oDVL.ToTable(True, "OU", "JV", "OcrCode3", "NewOU")
                                                For Each oddr As DataRow In oDT_CA.Rows
                                                    sSQL += "insert into [@AB_COSTALLOCATION] (Code, Name, U_Transid, U_PrcCode , U_OcrCode3, U_SourceOcrcode3)  values ((select isnull(max(cast(Code as numeric )),0) + 1 from [@AB_COSTALLOCATION]) ,(select isnull(max(cast(Code as numeric )),0) + 1 from [@AB_COSTALLOCATION])," & _
                                                    "'" & oddr("JV").ToString() & "','" & oddr("OU").ToString() & "', '" & oddr("OcrCode3").ToString() & "' , '" & oddr("NewOU").ToString() & "')"
                                                    ' sSQL += " update OPCH set  Indicator= 'CA' where DocEntry = '" & oddr("Base").ToString.Trim() & "' "
                                                    '  sSQL += " update OJDT set  Indicator= 'CA' where TransId = '" & oddr("JV").ToString.Trim() & "' "
                                                Next

                                                If sSQL.Length > 0 Then
                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Updating Flag " & sSQL, sFuncName)
                                                    orset.DoQuery(sSQL)
                                                End If
                                                ''If p_oDICompany.InTransaction Then
                                                ''    p_oDICompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                                ''End If
                                                ''If oDICompany(irow).InTransaction Then
                                                ''    oDICompany(irow).EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                                ''End If
                                            End If
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starts NON JV Group ", sFuncName)
                                            oDVL.RowFilter = "JVGroup<>'JV' and EntityCode='" & odrjv("EntityCode") & "'"
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("NON JV Group Count " & oDVL.Count, sFuncName)
                                            If oDVL.Count > 0 Then
                                                If p_oDICompany.InTransaction = False Then
                                                    p_oDICompany.StartTransaction()
                                                End If
                                                If oDICompany(irow).InTransaction = False Then
                                                    oDICompany(irow).StartTransaction()
                                                End If
                                                ''i_dSourceTable = New DataTable
                                                ''i_dSourceTable = oDVL.Table
                                                '' '' Dim dtGroup As DataTable = oDVL.ToTable(True, New String() {"JVGroup", "NewOU", "GL_NameT", "JV", "Base"})
                                                '' ''   Dim dtGroup As DataTable = oDVL.ToTable(True, New String() {i_sGroupByColumn, ii_sGroupByColumn, iii_sGroupByColumn, iiii_sGroupByColumn})
                                                ''Dim dtGroup As DataTable = oDVL.ToTable(True, "GL_NameT", "OU_BU_Budget", "Project", "NewOU", "Remarks", _
                                                ''                                         "TGL", "BU", "LOS", "Cat", "Base", "JV", "OcrCode3")
                                                ' ''adding column for the row count
                                                ''dtGroup.Columns.Add("Amount", GetType(Decimal))
                                                ' ''looping thru distinct values for the group, counting
                                                '' ''i_sGroupByColumn & " = '" & dr(i_sGroupByColumn) & "' AND " & ii_sGroupByColumn & " = '" & dr(ii_sGroupByColumn) & "' AND " & iii_sGroupByColumn & " = '" & dr(iii_sGroupByColumn) & "' AND " & iiii_sGroupByColumn & " = '" & dr(iiii_sGroupByColumn) & "'
                                                ''For Each dr As DataRow In dtGroup.Rows
                                                ''    dr("Amount") = i_dSourceTable.Compute("Sum(Amount)", " GL_NameT = '" & dr("GL_NameT") & "' AND OU_BU_Budget = '" & dr("OU_BU_Budget") & "' AND Project = '" & dr("Project") & "' " & _
                                                ''                                         " AND NewOU = '" & dr("NewOU") & "' AND Remarks = '" & dr("Remarks") & "'  AND Cat = '" & dr("Cat") & "' " & _
                                                ''                                        "  AND TGL = '" & dr("TGL") & "' AND JV = '" & dr("JV") & "' AND Base = '" & dr("Base") & "' AND" & _
                                                ''                                        "  BU = '" & dr("BU") & "' AND LOS = '" & dr("LOS") & "' AND OcrCode3 = '" & dr("OcrCode3") & "'")
                                                ''Next

                                                ''oDVJV = New DataView(dtGroup)

                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling JournalEntry_Posting_NONJV_Source()", sFuncName)
                                                sRef = String.Empty

                                                For Each oddr As DataRowView In oDVL
                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug1(oddr("GL_Code") & "  " & oddr("GL_NameT") & "  " & oddr("OU") & "  " & oddr("Amount") & "  " & oddr("TGL"), "JournalEntry_Posting_NONJV_Source")
                                                Next

                                                If JournalEntry_Posting_NONJV_Source(oDVL, p_oDICompany, sNow, sRef, sErrDesc) <> RTN_SUCCESS Then
                                                    If p_oDICompany.InTransaction Then
                                                        p_oDICompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                                    End If
                                                    For lCounter As Integer = 0 To UBound(oDICompany)
                                                        If Not oDICompany(lCounter) Is Nothing Then
                                                            If oDICompany(lCounter).Connected = True Then
                                                                If oDICompany(lCounter).InTransaction = True Then
                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback transaction on company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                    oDICompany(lCounter).EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                                                End If
                                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                oDICompany(lCounter).Disconnect()
                                                                oDICompany(lCounter) = Nothing
                                                            End If
                                                        End If
                                                    Next
                                                    Throw New ArgumentException(sErrDesc)
                                                End If
                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling JournalEntry_Posting_NONJV_Target()", sFuncName)
                                                ''i_dSourceTable = New DataTable
                                                ''i_dSourceTable = oDVL.ToTable
                                                '' '' Dim dtGroup As DataTable = oDVL.ToTable(True, New String() {"JVGroup", "NewOU", "GL_NameT", "JV", "Base"})
                                                '' ''   Dim dtGroup As DataTable = oDVL.ToTable(True, New String() {i_sGroupByColumn, ii_sGroupByColumn, iii_sGroupByColumn, iiii_sGroupByColumn})
                                                ''dtGroup = oDVL.ToTable(True, "GL_NameT", "OU_BU_Budget", "Project", "NewOU", "Remarks", _
                                                ''                                         "TGL", "BU", "LOS", "Cat")
                                                ' ''adding column for the row count
                                                ''dtGroup.Columns.Add("Amount", GetType(Decimal))
                                                ' ''looping thru distinct values for the group, counting
                                                '' ''i_sGroupByColumn & " = '" & dr(i_sGroupByColumn) & "' AND " & ii_sGroupByColumn & " = '" & dr(ii_sGroupByColumn) & "' AND " & iii_sGroupByColumn & " = '" & dr(iii_sGroupByColumn) & "' AND " & iiii_sGroupByColumn & " = '" & dr(iiii_sGroupByColumn) & "'
                                                ''For Each dr As DataRow In dtGroup.Rows
                                                ''    dr("Amount") = i_dSourceTable.Compute("Sum(Amount)", " GL_NameT = '" & dr("GL_NameT") & "' AND OU_BU_Budget = '" & dr("OU_BU_Budget") & "' AND Project = '" & dr("Project") & "' " & _
                                                ''                                         " AND NewOU = '" & dr("NewOU") & "' AND Remarks = '" & dr("Remarks") & "'  " & _
                                                ''                                        "  AND TGL = '" & dr("TGL") & "' AND " & _
                                                ''                                        "  BU = '" & dr("BU") & "' AND LOS = '" & dr("LOS") & "'  AND Cat = '" & dr("Cat") & "'")
                                                ''Next

                                                ''oDVJV = New DataView(dtGroup)

                                                For Each oddr As DataRowView In oDVL
                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug1(oddr("GL_Code") & "  " & oddr("GL_NameT") & "  " & oddr("OU") & "  " & oddr("Amount") & "  " & oddr("TGL"), "JournalEntry_Posting_NONJV_Target")
                                                Next

                                                If JournalEntry_Posting_NONJV_Target(oDVL, oDICompany(irow), sNow, sRef, sErrDesc) <> RTN_SUCCESS Then
                                                    If p_oDICompany.InTransaction Then
                                                        p_oDICompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                                    End If
                                                    For lCounter As Integer = 0 To UBound(oDICompany)
                                                        If Not oDICompany(lCounter) Is Nothing Then
                                                            If oDICompany(lCounter).Connected = True Then
                                                                If oDICompany(lCounter).InTransaction = True Then
                                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback transaction on company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                    oDICompany(lCounter).EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                                                End If
                                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                                oDICompany(lCounter).Disconnect()
                                                                oDICompany(lCounter) = Nothing
                                                            End If
                                                        End If
                                                    Next
                                                    Throw New ArgumentException("In Target :- " & sErrDesc)
                                                End If
                                                sSQL = ""
                                                Dim oDT_CA As DataTable = oDVL.ToTable(True, "OU", "JV", "OcrCode3", "NewOU")
                                                For Each oddr As DataRow In oDT_CA.Rows
                                                    sSQL += "insert into [@AB_COSTALLOCATION] (Code, Name, U_Transid, U_PrcCode , U_OcrCode3, U_SourceOcrcode3)  values ((select isnull(max(cast(Code as numeric )),0) + 1 from [@AB_COSTALLOCATION]) ,(select isnull(max(cast(Code as numeric )),0) + 1 from [@AB_COSTALLOCATION])," & _
                                                    "'" & oddr("JV").ToString() & "','" & oddr("OU").ToString() & "' , '" & oddr("OcrCode3").ToString() & "' , '" & oddr("NewOU").ToString() & "')"
                                                    ' sSQL += " update OPCH set  Indicator= 'CA' where DocEntry = '" & oddr("Base").ToString.Trim() & "' "
                                                    '  sSQL += " update OJDT set  Indicator= 'CA' where TransId = '" & oddr("JV").ToString.Trim() & "' "
                                                Next

                                                If sSQL.Length > 0 Then
                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("exporting into Cost Allocation Table " & sSQL, sFuncName)
                                                    orset.DoQuery(sSQL)
                                                End If
                                              
                                            End If
                                            irow += 1
                                        Next

                                        If p_oDICompany.InTransaction Then
                                            p_oDICompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                        End If
                                        For lCounter As Integer = 0 To UBound(oDICompany)
                                            If Not oDICompany(lCounter) Is Nothing Then
                                                If oDICompany(lCounter).Connected = True Then
                                                    If oDICompany(lCounter).InTransaction = True Then
                                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Commit transaction on company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                        oDICompany(lCounter).EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                                    End If
                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                    oDICompany(lCounter).Disconnect()
                                                    oDICompany(lCounter) = Nothing
                                                End If
                                            End If
                                        Next
                                        '' oForm.Items.Item("Item_21").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                        Dim oMAtric As SAPbouiCOM.Matrix = oForm.Items.Item("Item_24").Specific
                                        oMAtric.Clear()
                                        Dim oGridT2 As SAPbouiCOM.Grid = oForm.Items.Item("Item_23").Specific

                                        sSQL = "SELECT cast(Number as varchar(30)) 'SAP Invoice Number'  , REF2 'Vendor Invoice Number' , TAXDATE 'Vendor Invoice Date' ,REFDATE 'SAP Posting Date'  ,'JE' 'Vendor Name',MEMO 'Journal Remark' ,  [OcrCode3] 'Distribution Code', '' 'Purchasing Department', plan_id [Expenses]   from ( SELECT c.Number   , b.REF2  ,b.TAXDATE  ,c.REFDATE  ,c.MEMO  ,  b.[OcrCode3]  , a.AcctCode + ' - ' + a.AcctName [plan_id] , e.[PrcCode] [U_AB_NONPROJECT],    sum( round((ISNULL(b.Debit,0) - ISNULL(b.Credit,0)) /(d.[OcrTotal] /e.[PrcAmount] ),3)) AS total ,  sum(ISNULL(b.Debit,0) - ISNULL(b.Credit,0)) [Line]  FROM JDT1 b JOIN OACT a ON a.AcctCode = b.Account  JOIN OJDT c on b.[TransId] = c.[TransId]  left join OOCR d on d.OcrCode = b.OcrCode3  join OCR1 e on d.[OcrCode] = e.[OcrCode]  where  b.OcrCode3 in ('') and a.GroupMask = 5  and isnull(c.Indicator,'') <> 'CA' and 	b.Account >= '' and b.Account <= '' and c.transType IN ('0')  group by c.Number   , b.REF2  ,b.TAXDATE  ,c.REFDATE  ,c.MEMO  ,  b.[OcrCode3]  , a.AcctCode , a.AcctName , e.[PrcCode]    ) x "
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("SummaryReport show button click " & sSQL, sFuncName)
                                        Try
                                            oForm.DataSources.DataTables.Add("MyDataTable1")
                                        Catch ex As Exception
                                        End Try
                                        oGridT2.DataTable = oForm.DataSources.DataTables.Item("MyDataTable1")
                                        oForm.DataSources.DataTables.Item(1).ExecuteQuery(sSQL)
                                        oGridT2.CollapseLevel = 2
                                        oGridT2.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
                                        p_Summaryreport = String.Empty
                                        Threading.Thread.Sleep(1500)
                                        p_oSBOApplication.SetStatusBarMessage("Journal Entries Created Successfully ....!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Journal Entries Created Successfully ....!", sFuncName)

                                    Catch ex As Exception
                                        BubbleEvent = False
                                        p_oSBOApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                        sErrDesc = ex.Message
                                        If p_oDICompany.InTransaction Then
                                            p_oDICompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                        End If
                                        For lCounter As Integer = 0 To UBound(oDICompany)
                                            If Not oDICompany(lCounter) Is Nothing Then
                                                If oDICompany(lCounter).Connected = True Then
                                                    If oDICompany(lCounter).InTransaction = True Then
                                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback transaction on company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                        oDICompany(lCounter).EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                                    End If
                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting company database " & oDICompany(lCounter).CompanyDB, sFuncName)
                                                    oDICompany(lCounter).Disconnect()
                                                    oDICompany(lCounter) = Nothing
                                                End If
                                            End If
                                        Next
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                                        WriteToLogFile(Err.Description, sFuncName)
                                        ShowErr(sErrDesc)
                                    End Try
                                End If

                            End If


                        Case "GTF"
                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                                If pVal.ItemUID = "btnGntFile" Then
                                    Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                                    Dim oRset As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim sSQL As String = String.Empty
                                    Dim sCheck As String = String.Empty
                                    dtTable.Clear()

                                    Try
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling HeaderValidation()", sFuncName)

                                        SBO_Application.SetStatusBarMessage("Validation Process Started ........!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                        If HeaderValidation(oForm, sErrDesc) = 0 Then
                                            BubbleEvent = False
                                            Exit Sub
                                        End If

                                        If oForm.Items.Item("15").Specific.checked = True Then
                                            sCheck = "Y"
                                        Else
                                            sCheck = "N"
                                        End If


                                        If oDT_TxtFileGeneration.Rows.Count > 0 Then
                                            For imjs As Integer = 0 To oDT_TxtFileGeneration.Rows.Count - 1
                                                sSQL = "AE_SP001_TextFileGeneration " & _
                                                    "'" & oDT_TxtFileGeneration.Rows(imjs).Item("DateFrom").ToString & "', '" & oDT_TxtFileGeneration.Rows(imjs).Item("DateTo").ToString & "', " & _
                                                    "'" & oDT_TxtFileGeneration.Rows(imjs).Item("OUCodeFrom").ToString & "', '" & oDT_TxtFileGeneration.Rows(imjs).Item("OUCodeTo").ToString & "'," & _
                                                    "'" & oDT_TxtFileGeneration.Rows(imjs).Item("Entity").ToString & "'"
                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Sql Query " & sSQL, sFuncName)
                                                oRset.DoQuery(sSQL)
                                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConvertRecordset()", sFuncName)
                                                ConvertRecordset(oRset, sErrDesc)
                                            Next imjs
                                        End If


                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling HeaderValidation()", sFuncName)
                                        If Write_TextFile(dtTable, oDT_TxtFileGeneration.Rows(0).Item("FolderPath").ToString, sCheck, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                        p_oSBOApplication.StatusBar.SetText("File Generated Successfully ....... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS .......", sFuncName)


                                    Catch ex As Exception
                                        BubbleEvent = False
                                        sErrDesc = ex.Message
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                                        WriteToLogFile(Err.Description, sFuncName)
                                        ShowErr(sErrDesc)
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
                oCreationPackage.UniqueID = "FGTF"
                oCreationPackage.String = "Generate Text File"

                If Not p_oSBOApplication.Menus.Exists("FGTF") Then
                    oMenus.AddEx(oCreationPackage)
                End If

                oMenuItem = SBO_Application.Menus.Item("PWC")
                oMenus = oMenuItem.SubMenus

                'Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oCreationPackage.UniqueID = "BCA"
                oCreationPackage.String = "Cost Allocation"

                If Not p_oSBOApplication.Menus.Exists("BCA") Then
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
            Dim oForm As SAPbouiCOM.Form = Nothing

            If BusinessObjectInfo.ActionSuccess = True And BusinessObjectInfo.BeforeAction = False Then
                If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Then
                    Select Case BusinessObjectInfo.FormTypeEx
                        Case "392"
                            If bJEflag = True And p_sCancel.Length > 0 Then
                                oForm = p_oSBOApplication.Forms.GetFormByTypeAndCount(BusinessObjectInfo.FormTypeEx, iJEFormType)
                                Dim oRset As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Cancelling Documents " & p_sCancel, sFuncName)
                                oRset.DoQuery(p_sCancel)
                                p_sCancel = String.Empty
                                bJEflag = False
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Cancellation Completed with Success " & p_sCancel, sFuncName)
                            End If
                    End Select
                End If
            End If



        End Sub
    End Class
End Namespace



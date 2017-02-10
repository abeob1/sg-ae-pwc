Option Explicit On
Imports System.Xml
Imports System.IO
Imports System.Windows.Forms
Imports System.Globalization
Imports System.Net.Mail
Imports System.Configuration
Imports System.Reflection



Namespace AE_PWC_AO04

    Module modCommon



        Public Function ConnectDICompSSO(ByRef objCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
            ' ***********************************************************************************
            '   Function   :    ConnectDICompSSO()
            '   Purpose    :    Connect To DI Company Object
            '
            '   Parameters :    ByRef objCompany As SAPbobsCOM.Company
            '                       objCompany = set the SAP Company Object
            '                   ByRef sErrDesc As String
            '                       sErrDesc = Error Description to be returned to calling function
            '   Return     :    0 - FAILURE
            '                   1 - SUCCESS
            '   Author     :    Sri
            '   Date       :    29 April 2013
            '   Change     :
            ' ***********************************************************************************
            Dim sCookie As String = String.Empty
            Dim sConnStr As String = String.Empty
            Dim sFuncName As String = String.Empty
            Dim lRetval As Long
            Dim iErrCode As Int32
            Try
                sFuncName = "ConnectDICompSSO()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

                objCompany = New SAPbobsCOM.Company

                sCookie = objCompany.GetContextCookie
                sConnStr = p_oUICompany.GetConnectionContext(sCookie)
                'sConnStr = p_oSBOApplication.Company.GetConnectionContext(sCookie)
                lRetval = objCompany.SetSboLoginContext(sConnStr)

                If Not lRetval = 0 Then
                    Throw New ArgumentException("SetSboLoginContext of Single SignOn Failed.")
                End If
                p_oSBOApplication.StatusBar.SetText("Please Wait While Company Connecting... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                lRetval = objCompany.Connect
                If lRetval <> 0 Then
                    objCompany.GetLastError(iErrCode, sErrDesc)
                    Throw New ArgumentException("Connect of Single SignOn failed : " & sErrDesc)
                Else
                    p_oSBOApplication.StatusBar.SetText("Company Connection Has Established with the " & objCompany.CompanyName, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                End If
                ConnectDICompSSO = RTN_SUCCESS
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Catch exc As Exception
                sErrDesc = exc.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                ConnectDICompSSO = RTN_ERROR
            End Try
        End Function

        Public Function EntityLoad(ByRef oform As SAPbouiCOM.Form, ByRef sErrDesc As String) As Long

            Dim oMatrix As SAPbouiCOM.Matrix = oform.Items.Item("5").Specific
            Dim sSQL As String = String.Empty
            Dim oCombo As SAPbouiCOM.ComboBox = oform.Items.Item("Item_0").Specific
            Try
                sFuncName = "EntityLoad()"

                oform.Freeze(True)

                oMatrix.Columns.Item("Col_1").Visible = False
                oMatrix.Columns.Item("Col_2").Visible = False

                oCombo.ValidValues.Add("--Select--", "0")
                oCombo.ValidValues.Add("Chart of Account", "COA")
                oCombo.ValidValues.Add("Vendor / Customer", "BP")
                oCombo.ValidValues.Add("Item Master Data", "ITM")
                oCombo.ValidValues.Add("Dimension & Cost Center", "OPRC")
                oCombo.ValidValues.Add("Distribution Rule", "OOCR")
                oCombo.ValidValues.Add("Sales Person & Department", "OSLP")
                oCombo.ValidValues.Add("Users", "OUSR")
                oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)
                sSQL = "SELECT T0.[U_AB_COMCODE], T0.[U_AB_COMPANYNAME], T0.[U_AB_USERCODE], T0.[U_AB_PASSWORD]  FROM [dbo].[@AB_COMPANYDATA]  T0"

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("SQL " & sSQL, sFuncName)
                Try
                    oform.DataSources.DataTables.Add("U_AB_COMPANYNAME")
                Catch ex As Exception
                End Try
                oform.DataSources.DataTables.Item("U_AB_COMPANYNAME").ExecuteQuery(sSQL)
                oMatrix.Clear()
                'oMatrix.Columns.Item("V_1").DataBind.Bind("U_AB_COMPANYNAME", "Choose")
                oMatrix.Columns.Item("Col_0").DataBind.Bind("U_AB_COMPANYNAME", "U_AB_COMCODE")
                oMatrix.Columns.Item("V_0").DataBind.Bind("U_AB_COMPANYNAME", "U_AB_COMPANYNAME")
                oMatrix.Columns.Item("Col_1").DataBind.Bind("U_AB_COMPANYNAME", "U_AB_USERCODE")
                oMatrix.Columns.Item("Col_2").DataBind.Bind("U_AB_COMPANYNAME", "U_AB_PASSWORD")
                oMatrix.LoadFromDataSource()

                For imjs As Integer = 1 To oMatrix.RowCount
                    oMatrix.Columns.Item("V_-1").Cells.Item(imjs).Specific.String = imjs
                Next imjs


                'oMatrix.AutoResizeColumns()
                oMatrix.Columns.Item("V_0").Width = 200
                oMatrix.Columns.Item("Col_0").Width = 80
                oMatrix.Columns.Item("Col_3").Width = 100
                oMatrix.Columns.Item("Col_4").Width = 250

                oform.Freeze(False)
                EntityLoad = RTN_SUCCESS
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

            Catch ex As Exception
                oform.Freeze(False)
                sErrDesc = ex.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                EntityLoad = RTN_ERROR
            End Try
        End Function

        Public Function ConnectTargetDB(ByRef oTargetCmp As SAPbobsCOM.Company, _
                                        ByVal sTargetDB As String, _
                                        ByVal sSAPUser As String, _
                                        ByVal sSAPPwd As String, _
                                        ByRef sErrDesc As String) As Long
            ' **********************************************************************************
            'Function   :   ConnectTargetDB()
            'Purpose    :   Connect To Target Database
            '               This is for Intercompany Features
            '               
            'Parameters :   ByRef sErrDesc As String
            '                   sErrDesc=Error Description to be returned to calling function
            '               
            '                   =
            'Return     :   0 - FAILURE
            '               1 - SUCCESS
            'Author     :   Sri
            'Date       :   30 April 2013
            'Change     :
            ' **********************************************************************************

            Dim sFuncName As String = String.Empty
            Dim lRetval As Long
            Dim iErrCode As Integer
            Try
                sFuncName = "ConnectTargetDB()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

                oTargetCmp = Nothing
                oTargetCmp = New SAPbobsCOM.Company

                With oTargetCmp
                    .Server = p_oDICompany.Server                           'Name of the DB Server 
                    .DbServerType = p_oDICompany.DbServerType 'Database Type
                    .CompanyDB = sTargetDB                        'Enter the name of Target company
                    .UserName = sSAPUser                           'Enter the B1 user name
                    .Password = sSAPPwd                           'Enter the B1 password
                    .language = SAPbobsCOM.BoSuppLangs.ln_English          'Enter the logon language
                    .UseTrusted = False
                End With

                lRetval = oTargetCmp.Connect()
                If lRetval <> 0 Then
                    oTargetCmp.GetLastError(iErrCode, sErrDesc)
                    oTargetCmp.CompanyDB = sTargetDB                        'Enter the name of Target company
                    p_oSBOApplication.MessageBox("Connect to Target Company Failed :  " & sTargetDB & ". " & sErrDesc)
                    Throw New ArgumentException("Connect to Target Company Failed :  " & sTargetDB & ". " & sErrDesc)
                End If

                ConnectTargetDB = RTN_SUCCESS
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

            Catch exc As Exception
                ConnectTargetDB = RTN_ERROR
                sErrDesc = exc.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Finally

            End Try
        End Function

        Public Function AddButton(ByRef oForm As SAPbouiCOM.Form, _
                                  ByVal sButtonID As String, _
                                  ByVal sCaption As String, _
                                  ByVal sItemNo As String, _
                                  ByVal iSpacing As Integer, _
                                  ByVal iWidth As Integer, _
                                  ByVal blnVisable As Boolean, _
                                  ByRef sErrDesc As String, _
                                  Optional ByVal oType As SAPbouiCOM.BoButtonTypes = 0, _
                                  Optional ByVal sCFLObjType As String = "") As Long
            ' ***********************************************************************************
            '   Function   :    AddButton()
            '   Purpose    :    Add Button To Form
            '
            '   Parameters :    ByVal oForm As SAPbouiCOM.Form
            '                       oForm = set the SAP UI Form Object
            '                   ByVal sButtonID As String
            '                       sButtonID = Button UID
            '                   ByVal sCaption As String
            '                       sCaption = Caption
            '                   ByVal sItemNo As String
            '                       sItemNo = Next to Item UID
            '                   ByVal iSpacing As Integer
            '                       iSpacing = Spacing between sItemNo
            '                   ByVal iWidth As Integer
            '                       iWidth = Width
            '                   ByVal blnVisable As Boolean
            '                       blnVisible = True/False
            '                   ByRef sErrDesc As String
            '                       sErrDesc = Error Description to be returned to calling function
            '                   Optional ByVal oType As SAPbouiCOM.BoButtonTypes
            '                       oType = set the SAP UI Button Type Object
            '                   Optional ByVal sCFLObjType As String = ""
            '                       sCFLObjType = CFL Object Type
            '                   Optional ByVal sImgPath As String = ""
            '                       sImgPath = Image Path
            '   Return     :    0 - FAILURE
            '                   1 - SUCCESS
            '   Author     :    Jason Ham
            '   Date       :    9 Jan 2007
            '   Change     :
            '                   9 Jan 2008 (Jason) Add Object Link
            ' ***********************************************************************************
            Dim oItems As SAPbouiCOM.Items
            Dim oItem As SAPbouiCOM.Item
            Dim oButton As SAPbouiCOM.Button
            Dim sFuncName As String = String.Empty

            Try
                sFuncName = "AddButton()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

                oItems = oForm.Items
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Add BUTTON Item", sFuncName)
                oItem = oItems.Add(sButtonID, SAPbouiCOM.BoFormItemTypes.it_BUTTON)
                If sCaption <> "" Then
                    oItem.Specific.Caption = sCaption
                End If
                oItem.Visible = blnVisable
                oItem.Left = oItems.Item(sItemNo).Left + oItems.Item(sItemNo).Width + iSpacing
                oItem.Height = oItems.Item(sItemNo).Height
                oItem.Top = oItems.Item(sItemNo).Top
                oItem.Width = iWidth
                oButton = oItem.Specific
                oButton.Type = oType    'default is Caption type.

                If oType = 1 Then oButton.Image = "CHOOSE_ICON" 'This line will fire if the button type is image

                If sCFLObjType <> "" Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Add User Data Source :" & sButtonID, sFuncName)
                    oForm.DataSources.UserDataSources.Add(sButtonID, SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("AddChooseFromList" & sButtonID, sFuncName)
                    AddChooseFromList(oForm, sCFLObjType, sButtonID, sErrDesc)
                    oButton.ChooseFromListUID = sButtonID
                End If

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
                AddButton = RTN_SUCCESS
            Catch exc As Exception
                AddButton = RTN_ERROR
                sErrDesc = exc.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Finally
                oItems = Nothing
                oItem = Nothing
            End Try

        End Function

        Public Function AddChooseFromList(ByVal oForm As SAPbouiCOM.Form, ByVal sCFLObjType As String, ByVal sItemUID As String, ByRef sErrDesc As String) As Long
            ' ***********************************************************************************
            '   Function   :    AddChooseFromList()
            '   Purpose    :    Create Choose From List For User Define Form
            '
            '   Parameters :    ByVal oForm As SAPbouiCOM.Form
            '                       oForm = set the SAP UI Form Object
            '                   ByVal sCFLObjType As String
            '                       sCFLObjType = set SAP UI Choose From List Object Type
            '                   ByRef sErrDesc As String
            '                       sErrDesc = Error Description to be returned to calling function
            '   Return     :    0 - FAILURE
            '                   1 - SUCCESS
            '   Author     :    Jason Ham
            '   Date       :    30/12/2007
            '   Change     :
            ' ***********************************************************************************
            Dim sFuncName As String = String.Empty
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams

            Try

                sFuncName = "AddChooseFromList"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating 'ChooseFromLists' and 'cot_ChooseFromListCreationParams' objects", sFuncName)
                oCFLs = oForm.ChooseFromLists
                oCFLCreationParams = p_oSBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Setting Choose From List Parameter properties", sFuncName)
                'Only Single Selection
                oCFLCreationParams.MultiSelection = False
                'Determine the Object Type
                oCFLCreationParams.ObjectType = sCFLObjType
                'Item UID as Unique ID for CFL
                oCFLCreationParams.UniqueID = sItemUID

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Choose From List Parameter", sFuncName)
                oCFL = oCFLs.Add(oCFLCreationParams)

                AddChooseFromList = RTN_SUCCESS
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

            Catch exc As Exception
                AddChooseFromList = RTN_ERROR
                sErrDesc = exc.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            End Try

        End Function

        Public Function AddUserDataSrc(ByRef oForm As SAPbouiCOM.Form, ByVal sDSUID As String, _
                                       ByRef sErrDesc As String, ByVal oDataType As SAPbouiCOM.BoDataType, _
                                       Optional ByVal lLen As Long = 0) As Long
            ' ***********************************************************************************
            '   Function   :    AddUserDataSrc()
            '   Purpose    :    Add User Data Source
            '
            '   Parameters :    ByVal oForm As SAPbouiCOM.Form
            '                       oForm = set the SAP UI Form Object
            '                   ByVal sDSUID As String
            '                       sDSUID = Data Set UID
            '                   ByRef sErrDesc As String
            '                       sErrDesc = Error Description to be returned to calling function
            '                   ByVal oDataType As SAPbouiCOM.BoDataType
            '                       oDataType = set the SAP UI BoDataType Object
            '                   Optional ByVal lLen As Long = 0
            '                       lLen= Length
            '   Return     :    0 - FAILURE
            '                   1 - SUCCESS
            '   Author     :    Dev
            '   Date       :    23 Jan 2007
            '   Change     :
            ' ***********************************************************************************
            Dim sFuncName As String = String.Empty
            Try
                sFuncName = "AddUserDataSrc()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

                If lLen = 0 Then
                    oForm.DataSources.UserDataSources.Add(sDSUID, oDataType)
                Else
                    oForm.DataSources.UserDataSources.Add(sDSUID, oDataType, lLen)
                End If

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
                AddUserDataSrc = RTN_SUCCESS
            Catch exc As Exception
                sErrDesc = exc.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                AddUserDataSrc = RTN_ERROR
            End Try
        End Function

        Public Function AddItem(ByRef oForm As SAPbouiCOM.Form, ByVal sItemUID As String, ByVal bEnable As Boolean, _
                                ByVal oItemType As SAPbouiCOM.BoFormItemTypes, ByRef sErrDesc As String, _
                                Optional ByVal sCaption As String = "", Optional ByVal iPos As Integer = 0, _
                                Optional ByVal sPosItemUID As String = "", Optional ByVal lSpace As Long = 5, _
                                Optional ByVal lLeft As Long = 0, Optional ByVal lTop As Long = 0, _
                                Optional ByVal lHeight As Long = 0, Optional ByVal lWidth As Long = 0, _
                                Optional ByVal lFromPane As Long = 0, Optional ByVal lToPane As Long = 0, _
                                Optional ByVal sCFLObjType As String = "", Optional ByVal sCFLAlias As String = "", _
                                Optional ByVal oLinkedObj As SAPbouiCOM.BoLinkedObject = 0, _
                                Optional ByVal sBindTbl As String = "", Optional ByVal sAlias As String = "", _
                                Optional ByVal bDisplayDesc As Boolean = False) As Long
            ' ***********************************************************************************
            '   Function   :    AddItem()
            '   Purpose    :    Add Form's Item
            '
            '   Parameters :    ByVal oForm As SAPbouiCOM.Form
            '                       oForm = set the SAP UI Form Type
            '                   ByVal sItemUID As String
            '                       sItemUID = Item's ID
            '                   ByVal bEnable As Boolean
            '                       bEnable = Enable or Disable The Item
            '                   ByVal oItemType As SAPbouiCOM.BoFormItemTypes
            '                       oItemType = Item's Type
            '                   ByRef sErrDesc As String
            '                       sErrDesc = Error Description to be returned to calling function
            '                   Optional ByVal sCaption As String = ""
            '                       sCaption = Caption
            '                   Optional ByVal iPos As Integer = 0
            '                       iPos = Position.
            '                           Case 1 Left os sPosItemUID
            '                           Case 2 Right of sPosItemUID
            '                           Case 3 Top of sPosItemUID
            '                           Case Else Below sPosItemUID
            '                   Optional ByVal sPosItemUID As String = ""
            '                       sPosItemUID=Returns or sets the beginning of range specifying on which panes the item is visible. 0 by default
            '                   Optional ByVal lSpace As Long = 5
            '                       lSpace=sets the item space between oItem and sPosItemUID
            '                   Optional ByVal lLeft As Long = 0
            '                       lLeft=sets the item Left.
            '                   Optional ByVal lTop As Long = 0
            '                       lTop=sets the item top.
            '                   Optional ByVal lHeight As Long = 0
            '                       lHeight=sets the item height.
            '                   Optional ByVal lWidth As Long = 0
            '                       lWidth=sets the item weight.
            '                   Optional ByVal lFromPane As Long = 0
            '                       lFromPane=sets the beginning of range specifying on which panes the item is visible. 0 by default.
            '                   Optional ByVal lToPane As Long = 0
            '                       lToPane=sets the beginning of range specifying on which panes the item is visible. 0 by default.
            '                   Optional ByVal sCFLObjType As String = ""
            '                       sCFLObjType=CFL Obj Type
            '                   Optional ByVal sCFLAlias As String = ""
            '                       sCFLAlias=CFL Alias
            '                   Optional ByVal sBindTbl As String = ""
            '                       sBindTbl=Bind Table 
            '                   Optional ByVal sAlias As String = ""
            '                       sAlias=Alias
            '                   Optional ByVal bDisplayDesc As Boolean = False
            '                       bDisplayDesc=Returns or sets a a boolean value specifying whether or not to show the description of valid values of a ComboBox item. 
            '                                   True - displays the description of the valid value.
            '                                   False - displays the value of the selected valid value. 
            '   Return     :    0 - FAILURE
            '                   1 - SUCCESS
            '   Author     :    Sri
            '   Date       :    29/04/2013
            ' ***********************************************************************************

            Dim oItem As SAPbouiCOM.Item
            Dim oPosItem As SAPbouiCOM.Item
            Dim oEdit As SAPbouiCOM.EditText
            Dim sFuncName As String = String.Empty

            Try
                sFuncName = "AddItem()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function. Item: " & sItemUID, sFuncName)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding item", sFuncName)
                oItem = oForm.Items.Add(sItemUID, oItemType)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Setting item properties", sFuncName)
                If Trim(sPosItemUID) <> "" Then
                    oPosItem = oForm.Items.Item(sPosItemUID)
                    oItem.Enabled = bEnable
                    oItem.Height = oPosItem.Height
                    oItem.Width = oPosItem.Width
                    Select Case iPos
                        Case 1      'Left of sPosItemUID
                            oItem.Left = oPosItem.Left - lSpace
                            oItem.Top = oPosItem.Top
                        Case 2      '2=Right of sPosItemUID
                            oItem.Left = oPosItem.Left + oPosItem.Width + lSpace
                            oItem.Top = oPosItem.Top
                        Case 3      '3=Top of sPosItemUID
                            oItem.Left = oPosItem.Left
                            oItem.Top = oPosItem.Top - lSpace
                        Case 4
                            oItem.Left = oPosItem.Left + lSpace
                            oItem.Top = oPosItem.Top + lSpace
                        Case Else   'Below sPosItemUID
                            oItem.Left = oPosItem.Left
                            oItem.Top = oPosItem.Top + oPosItem.Height + lSpace
                    End Select
                End If

                If lTop <> 0 Then oItem.Top = lTop
                If lLeft <> 0 Then oItem.Left = lLeft
                If lHeight <> 0 Then oItem.Height = lHeight
                If lWidth <> 0 Then oItem.Width = lWidth

                If Trim(sBindTbl) <> "" Or Trim(sAlias) <> "" Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding item DataSource", sFuncName)
                    oItem.Specific.DataBind.SetBound(True, sBindTbl, sAlias)
                End If

                oItem.FromPane = lFromPane
                oItem.ToPane = lToPane
                oItem.DisplayDesc = bDisplayDesc

                If Trim(sCaption) <> "" Then oItem.Specific.Caption = sCaption

                If sCFLObjType <> "" And oItem.Type = SAPbouiCOM.BoFormItemTypes.it_EDIT Then
                    'If Choose From List Item
                    oForm.DataSources.UserDataSources.Add(sItemUID, SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddChooseFromList()", sFuncName)
                    AddChooseFromList(oForm, sCFLObjType, sItemUID, sErrDesc)
                    oEdit = oItem.Specific
                    oEdit.DataBind.SetBound(True, "", sItemUID)
                    oEdit.ChooseFromListUID = sItemUID
                    oEdit.ChooseFromListAlias = sCFLAlias
                End If

                If oLinkedObj <> 0 Then
                    Dim oLink As SAPbouiCOM.LinkedButton
                    oItem.LinkTo = sPosItemUID 'ID of the edit text used to idenfity the object to open
                    oLink = oItem.Specific
                    oLink.LinkedObject = oLinkedObj
                End If
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
                AddItem = RTN_SUCCESS
            Catch exc As Exception
                sErrDesc = exc.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                AddItem = RTN_ERROR
            Finally
                oItem = Nothing
                oPosItem = Nothing
                GC.Collect()
            End Try
        End Function

        Public Function StartTransaction(ByRef sErrDesc As String) As Long
            ' ***********************************************************************************
            '   Function   :    StartTransaction()
            '   Purpose    :    Start DI Company Transaction
            '
            '   Parameters :    ByRef sErrDesc As String
            '                       sErrDesc = Error Description to be returned to calling function
            '   Return     :   0 - FAILURE
            '                   1 - SUCCESS
            '   Author     :   Sri
            '   Date       :   29 April 2013
            '   Change     :
            ' ***********************************************************************************
            Dim sFuncName As String = String.Empty

            Try
                sFuncName = "StartTransaction()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

                If p_oDICompany.InTransaction Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Found hanging transaction.Rolling it back.", sFuncName)
                    p_oDICompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                End If

                p_oDICompany.StartTransaction()

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
                StartTransaction = RTN_SUCCESS
            Catch exc As Exception
                sErrDesc = exc.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                StartTransaction = RTN_ERROR
            End Try

        End Function

        Public Function RollBackTransaction(ByRef sErrDesc As String) As Long
            ' ***********************************************************************************
            '   Function   :    RollBackTransaction()
            '   Purpose    :    Roll Back DI Company Transaction
            '
            '   Parameters :    ByRef sErrDesc As String
            '                       sErrDesc = Error Description to be returned to calling function
            '   Return     :    0 - FAILURE
            '                   1 - SUCCESS
            '   Author     :    Sri
            '   Date       :    29 April 2013
            '   Change     :
            ' ***********************************************************************************
            Dim sFuncName As String = String.Empty

            Try
                sFuncName = "RollBackTransaction()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

                If p_oDICompany.InTransaction Then
                    p_oDICompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                Else
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No active transaction found for rollback", sFuncName)
                End If

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
                RollBackTransaction = RTN_SUCCESS
            Catch exc As Exception
                sErrDesc = exc.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                RollBackTransaction = RTN_ERROR
            Finally
                GC.Collect()
            End Try

        End Function

        Public Function fillopen() As String
            Dim sFuncName As String = String.Empty
            Dim ShowFolderBrowserThread As Threading.Thread
            Try
                sFuncName = "fillopen()"
                ' If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)
                ShowFolderBrowserThread = New Threading.Thread(AddressOf ShowFileDialog)
                If ShowFolderBrowserThread.ThreadState = System.Threading.ThreadState.Unstarted Then
                    ShowFolderBrowserThread.SetApartmentState(System.Threading.ApartmentState.STA)
                    ShowFolderBrowserThread.Start()
                    ShowFolderBrowserThread.Join()
                ElseIf ShowFolderBrowserThread.ThreadState = System.Threading.ThreadState.Stopped Then
                    ShowFolderBrowserThread.Start()
                    ShowFolderBrowserThread.Join()

                End If
                While ShowFolderBrowserThread.ThreadState = Threading.ThreadState.Running
                    Windows.Forms.Application.DoEvents()
                End While

                If Not String.IsNullOrEmpty(p_sSelectedFilepath) Then
                    Return p_sSelectedFilepath
                End If

                ' If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Catch ex As Exception
                fillopen = String.Empty
                sErrDesc = ex.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            End Try
        End Function

        Public Sub ShowFileDialog()

            Dim oDialogBox As New OpenFileDialog
            Dim sMdbFilePath As String
            Dim sFuncName As String = String.Empty
            '' Dim oProcesses() As System.Diagnostics.Process
            Try

                sFuncName = "ShowFileDialog()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)


                Dim OpenFilePath As Browse
                OpenFilePath = New Browse
                OpenFilePath.Show()
                OpenFilePath.Visible = False
                OpenFilePath.TopMost = True
                OpenFilePath.OpenFileDialog1.Multiselect = False
                OpenFilePath.OpenFileDialog1.Title = "SAP Business One"
                OpenFilePath.OpenFileDialog1.InitialDirectory = "C:\"
                OpenFilePath.OpenFileDialog1.FileName = String.Empty

                If OpenFilePath.OpenFileDialog1.ShowDialog = DialogResult.OK Then
                    sMdbFilePath = OpenFilePath.OpenFileDialog1.FileName
                    p_sSelectedFilepath = OpenFilePath.OpenFileDialog1.FileName
                    ''  p_sSelectedFileName = OpenFilePath.OpenFileDialog1.SafeFileName
                    OpenFilePath.OpenFileDialog1.Dispose()
                    OpenFilePath.Close()

                Else
                    p_sSelectedFilepath = String.Empty
                    OpenFilePath.Close()
                    System.Windows.Forms.Application.ExitThread()
                End If


                ' ''oDialogBox.Multiselect = False
                ' ''oDialogBox.Filter = "All files (*.*)|*.*|All files (*.*)|*.*" '"All files(*.CSV)|*.CSV"
                ' ''Dim filterindex As Integer = 0
                ' ''Try
                ' ''    filterindex = 0
                ' ''Catch ex As Exception
                ' ''End Try

                ' ''oDialogBox.FilterIndex = filterindex
                ' ''oDialogBox.InitialDirectory = "C:\" ' orset.Fields.Item("C:\").Value
                ' ''oDialogBox.RestoreDirectory = True

                ' ''oProcesses = System.Diagnostics.Process.GetProcessesByName("SAP Business One")
                '' ''If oProcesses.Length <> 0 Then
                '' ''    For i As Integer = 0 To oProcesses.Length - 1
                ' ''Dim MyWindow As New clsWindowWrapper(oProcesses(0).MainWindowHandle)
                ' ''Dim ret As DialogResult = oDialogBox.ShowDialog(MyWindow)
                ' ''If ret = DialogResult.OK Then
                ' ''    sMdbFilePath = oDialogBox.FileName
                ' ''    p_sSelectedFilepath = oDialogBox.FileName
                ' ''    p_sSelectedFileName = oDialogBox.SafeFileName
                ' ''    oDialogBox.Dispose()
                ' ''Else
                ' ''    p_sSelectedFilepath = String.Empty
                ' ''    System.Windows.Forms.Application.ExitThread()
                ' ''End If
                '    Next
                'End If

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Catch ex As Exception
                p_sSelectedFilepath = String.Empty
                sErrDesc = ex.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Finally
            End Try
        End Sub

        Public Function TransactionLog(ByVal sBpCode As String, ByVal sBpName As String, ByVal dDate As Date, ByVal dBalance As Double, ByVal sEmailAdd As String, _
                                 ByVal sStatus As String, ByVal sErrorMsg As String, ByRef oCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long

            Dim sFuncName As String = String.Empty
            Dim Ret As Integer
            Dim str As String
            Dim oRset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Try
                sFuncName = "ErrorLog()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function...", sFuncName)
                Dim oUserTable As SAPbobsCOM.UserTable
                oRset.DoQuery("SELECT (max(convert(numeric,isnull(T0.[Code],0))) + 1) [Code] FROM [dbo].[@AE_ESOALOG]  T0")
                oUserTable = oCompany.UserTables.Item("AE_ESOALOG")
                ' oUserTable.GetByKey("@AE_AGINGLOG")
                'Set default, mandatory fields
                oUserTable.Code = oRset.Fields.Item("Code").Value
                oUserTable.Name = oRset.Fields.Item("Code").Value

                oUserTable.UserFields.Fields.Item("U_BPCode").Value = sBpCode
                oUserTable.UserFields.Fields.Item("U_BPName").Value = sBpName
                oUserTable.UserFields.Fields.Item("U_Soa_Date").Value = dDate
                oUserTable.UserFields.Fields.Item("U_Balance").Value = dBalance
                oUserTable.UserFields.Fields.Item("U_EmailID").Value = sEmailAdd
                oUserTable.UserFields.Fields.Item("U_Status").Value = sStatus
                oUserTable.UserFields.Fields.Item("U_ErrMsg").Value = sErrorMsg
                oUserTable.UserFields.Fields.Item("U_EDate").Value = oCompany.GetDBServerDate()
                oUserTable.UserFields.Fields.Item("U_user").Value = oCompany.UserName
                oUserTable.Add()
                oCompany.GetLastError(Ret, str)

                If Ret <> 0 Then
                    oCompany.GetLastError(Ret, str)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & str, sFuncName)
                    WriteToLogFile(str, sFuncName)
                Else
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Data Added successfuly", sFuncName)
                End If
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
                TransactionLog = RTN_SUCCESS
            Catch ex As Exception
                WriteToLogFile(ex.Message, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                TransactionLog = RTN_ERROR
            End Try

        End Function

        Public Function CommitTransaction(ByRef sErrDesc As String) As Long
            ' ***********************************************************************************
            '   Function   :    CommitTransaction()
            '   Purpose    :    Commit DI Company Transaction
            '
            '   Parameters :    ByRef sErrDesc As String
            '                       sErrDesc=Error Description to be returned to calling function
            '   Return     :    0 - FAILURE
            '                   1 - SUCCESS
            '   Author     :    Sri
            '   Date       :    29 April 2013
            '   Change     :
            ' ***********************************************************************************
            Dim sFuncName As String = String.Empty

            Try
                sFuncName = "CommitTransaction()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

                If p_oDICompany.InTransaction Then
                    p_oDICompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                Else
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No active transaction found for commit", sFuncName)
                End If

                CommitTransaction = RTN_SUCCESS
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Catch exc As Exception
                CommitTransaction = RTN_ERROR
                sErrDesc = exc.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            End Try

        End Function

        Public Function DisplayStatus(ByVal oFrmParent As SAPbouiCOM.Form, ByVal sMsg As String, ByRef sErrDesc As String) As Long
            ' ***********************************************************************************
            '   Function   :    DisplayStatus()
            '   Purpose    :    Display Status Message while loading 
            '
            '   Parameters :    ByVal oFrmParent As SAPbouiCOM.Form
            '                       oFrmParent = set the SAP UI Form Object
            '                   ByVal sMsg As String
            '                       sMsg = set the Display Message information
            '                   ByRef sErrDesc As String
            '                       sErrDesc = Error Description to be returned to calling function
            '   Return     :    0 - FAILURE
            '                   1 - SUCCESS
            '   Author     :    Sri
            '   Date       :   29 April 2013
            '   Change     :
            ' ***********************************************************************************
            Dim oForm As SAPbouiCOM.Form
            Dim oItem As SAPbouiCOM.Item
            Dim oTxt As SAPbouiCOM.StaticText
            Dim creationPackage As SAPbouiCOM.FormCreationParams
            Dim iCount As Integer
            Dim sFuncName As String = String.Empty

            Try
                sFuncName = "DisplayStatus"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)
                'Check whether the form exists.If exists then close the form
                For iCount = 0 To p_oSBOApplication.Forms.Count - 1
                    oForm = p_oSBOApplication.Forms.Item(iCount)
                    If oForm.UniqueID = "dStatus" Then
                        oForm.Close()
                        Exit For
                    End If
                Next iCount
                'Add Form
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating form Assign Department", sFuncName)
                creationPackage = p_oSBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
                creationPackage.UniqueID = "dStatus"
                creationPackage.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_FixedNoTitle
                creationPackage.FormType = "AB_dStatus"
                oForm = p_oSBOApplication.Forms.AddEx(creationPackage)
                With oForm
                    .AutoManaged = False
                    .Width = 300
                    .Height = 100
                    If oFrmParent Is Nothing Then
                        .Left = (Screen.PrimaryScreen.WorkingArea.Width - oForm.Width) / 2
                        .Top = (Screen.PrimaryScreen.WorkingArea.Height - oForm.Height) / 2.5
                    Else
                        .Left = ((oFrmParent.Left * 2) + oFrmParent.Width - oForm.Width) / 2
                        .Top = ((oFrmParent.Top * 2) + oFrmParent.Height - oForm.Height) / 2
                    End If
                End With

                'Add Label
                oItem = oForm.Items.Add("3", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                oItem.Top = 40
                oItem.Left = 40
                oItem.Width = 250
                oTxt = oItem.Specific
                oTxt.Caption = sMsg
                oForm.Visible = True

                DisplayStatus = RTN_SUCCESS
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Catch exc As Exception
                DisplayStatus = RTN_ERROR
                sErrDesc = exc.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Finally
                creationPackage = Nothing
                oForm = Nothing
                oItem = Nothing
                oTxt = Nothing
            End Try

        End Function

        Public Function EndStatus(ByRef sErrDesc As String) As Long
            ' ***********************************************************************************
            '   Function   :    EndStatus()
            '   Purpose    :    Close Status Window
            '
            '   Parameters :    ByRef sErrDesc As String
            '                       sErrDesc = Error Description to be returned to calling function
            '   Return     :    0 - FAILURE
            '                   1 - SUCCESS
            '   Author     :    Sri
            '   Date       :    29 April 2013
            '   Change     :
            ' ***********************************************************************************
            Dim oForm As SAPbouiCOM.Form
            Dim iCount As Integer
            Dim sFuncName As String = String.Empty

            Try
                sFuncName = "EndStatus()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)
                'Check whether the form is exist. If exist then close the form
                For iCount = 0 To p_oSBOApplication.Forms.Count - 1
                    oForm = p_oSBOApplication.Forms.Item(iCount)
                    If oForm.UniqueID = "dStatus" Then
                        oForm.Close()
                        Exit For
                    End If
                Next iCount
                EndStatus = RTN_SUCCESS
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Catch exc As Exception
                EndStatus = RTN_ERROR
                sErrDesc = exc.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Finally
                oForm = Nothing
            End Try

        End Function

        Public Sub ShowErr(ByVal sErrMsg As String)
            ' ***********************************************************************************
            '   Function   :    ShowErr()
            '   Purpose    :    Show Error Message
            '   Parameters :  
            '                   ByVal sErrDesc As String
            '                       sErrDesc = Error Description to be returned to calling function
            '   Return     :    0 - FAILURE
            '                   1 - SUCCESS
            '   Author     :    Dev
            '   Date       :    23 Jan 2007
            '   Change     :
            ' ***********************************************************************************
            Try
                If sErrMsg <> "" Then
                    If Not p_oSBOApplication Is Nothing Then
                        If p_iErrDispMethod = ERR_DISPLAY_STATUS Then

                            p_oSBOApplication.SetStatusBarMessage("Error : " & sErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short)
                        ElseIf p_iErrDispMethod = ERR_DISPLAY_DIALOGUE Then
                            p_oSBOApplication.MessageBox("Error : " & sErrMsg)
                        End If
                    End If
                End If
            Catch exc As Exception
                WriteToLogFile(exc.Message, "ShowErr()")
            End Try
        End Sub

        Public Sub UpdateXML(ByVal oDICompany As SAPbobsCOM.Company, ByVal oDITargetComp As SAPbobsCOM.Company, _
                                 ByVal sNode As String, ByVal sTblName As String, ByVal sField1 As String, ByVal sField2 As String, _
                                 ByVal bIsNumeric As Boolean, ByRef oXMLDoc As XmlDocument, ByRef sXMLFile As String)

            Dim oNode As XmlNode
            Dim sFuncName As String = String.Empty
            Dim sSQL As String = String.Empty
            Dim oRs As SAPbobsCOM.Recordset
            Dim iCode As Integer
            Dim sCode As String = String.Empty

            Try
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Updating " & sField1 & " in XML file..", sFuncName)
                oRs = oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                oNode = oXMLDoc.SelectSingleNode(sNode)

                If Not IsNothing(oNode) Then
                    If Not oNode.InnerText = String.Empty Then
                        If bIsNumeric Then
                            iCode = CInt(oNode.InnerText)

                            If sTblName = "OLGT" Then
                                If CInt(oNode.InnerText) = 0 Then iCode = 1
                            End If


                            sSQL = " SELECT " & sField1 & " from  [" & oDITargetComp.CompanyDB.ToString & "].[dbo]." & sTblName & _
                                   " WHERE " & sField2 & " in (select " & sField2 & " from " & sTblName & " WHERE " & sField1 & "=" & iCode & ")"
                        Else
                            sCode = oNode.InnerText
                            sSQL = " SELECT " & sField1 & " from  [" & oDITargetComp.CompanyDB.ToString & "].[dbo]." & sTblName & _
                                   " WHERE " & sField2 & " in (select " & sField2 & " from " & sTblName & " WHERE " & sField1 & "='" & sCode & "')"
                        End If

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL Query" & sSQL, sFuncName)
                        oRs.DoQuery(sSQL)
                        If Not oRs.EoF Then
                            oNode.InnerText = oRs.Fields.Item(0).Value
                        Else
                            oNode.ParentNode.RemoveChild(oNode)
                            oXMLDoc.Save(sXMLFile)
                        End If
                        oXMLDoc.Save(sXMLFile)
                    Else
                        oNode.ParentNode.RemoveChild(oNode)
                        oXMLDoc.Save(sXMLFile)
                    End If
                End If

            Catch ex As Exception

            End Try

        End Sub

        Public Sub LoadFromXML(ByVal FileName As String, ByVal Sbo_application As SAPbouiCOM.Application)
            Try
                Dim oXmlDoc As New Xml.XmlDocument
                Dim sPath As String
                ''sPath = IO.Directory.GetParent(Application.StartupPath).ToString
                sPath = Application.StartupPath.ToString
                'oXmlDoc.Load(sPath & "\AE_FleetMangement\" & FileName)
                oXmlDoc.Load(sPath & "\" & FileName)
                ' MsgBox(Application.StartupPath)

                Sbo_application.LoadBatchActions(oXmlDoc.InnerXml)
            Catch ex As Exception
                MsgBox(ex)
            End Try

        End Sub

        Function HeaderValidation(FormUID As SAPbouiCOM.Form, ByRef sErrDesc As String) As Long
            Dim sFuncName As String = String.Empty
            Dim oMatrix As SAPbouiCOM.Matrix = FormUID.Items.Item("5").Specific
            Dim oCheckbox As SAPbouiCOM.CheckBox
            Dim oFlag As Boolean = False
            oDT_Entities = New DataTable

            oDT_Entities.Columns.Add("Sno", GetType(String))
            oDT_Entities.Columns.Add("Entity", GetType(String))
            oDT_Entities.Columns.Add("EntityDesc", GetType(String))
            oDT_Entities.Columns.Add("UserName", GetType(String))
            oDT_Entities.Columns.Add("Password", GetType(String))

            Try
                sFuncName = "HeaderValidation()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

                If FormUID.Items.Item("Item_0").Specific.value.ToString.Trim() = "--Select--" Then
                    p_oSBOApplication.StatusBar.SetText("Master Data Type can`t be Empty ....... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    FormUID.ActiveItem = "Item_0"
                    Return RTN_ERROR

                ElseIf FormUID.Items.Item("txtCode").Specific.string = String.Empty Then
                    p_oSBOApplication.StatusBar.SetText("Sync Code can`t be Empty ....... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    FormUID.ActiveItem = "txtCode"
                    Return RTN_ERROR
                End If

                If FormUID.Items.Item("Item_0").Specific.selected.description = "OSLP" Then
                    If FormUID.Items.Item("Item_3").Specific.checked = False And FormUID.Items.Item("Item_4").Specific.checked = False Then
                        p_oSBOApplication.StatusBar.SetText("Check Either Sales employee / Department ....... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        FormUID.ActiveItem = "txtCode"
                        Return RTN_ERROR
                    ElseIf FormUID.Items.Item("Item_3").Specific.checked = True And FormUID.Items.Item("Item_4").Specific.checked = True Then
                        p_oSBOApplication.StatusBar.SetText("Check Either Sales employee / Department, can`t both at same time  ....... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        FormUID.ActiveItem = "txtCode"
                        Return RTN_ERROR
                    End If
                End If

                For imjs As Integer = 1 To oMatrix.RowCount
                    oCheckbox = oMatrix.Columns.Item("V_1").Cells.Item(imjs).Specific
                    If oCheckbox.Checked = True Then
                        oDT_Entities.Rows.Add(imjs, oMatrix.Columns.Item("Col_0").Cells.Item(imjs).Specific.String, oMatrix.Columns.Item("V_0").Cells.Item(imjs).Specific.String, oMatrix.Columns.Item("Col_1").Cells.Item(imjs).Specific.String, oMatrix.Columns.Item("Col_2").Cells.Item(imjs).Specific.String)
                    End If

                Next imjs

                If oDT_Entities.Rows.Count = 0 Then
                    p_oSBOApplication.StatusBar.SetText("Please choose Entity ....... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Return RTN_ERROR
                End If

                HeaderValidation = RTN_SUCCESS
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

            Catch ex As Exception
                p_oSBOApplication.StatusBar.SetText("HeadValidation Function : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                sErrDesc = ex.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                HeaderValidation = RTN_ERROR
            End Try
            Return RTN_SUCCESS
        End Function

        Public Function GetDatasetFromExcel(ByRef oForm As SAPbouiCOM.Form, ByVal CurrFileToUpload As String, ByRef sErrDesc As String) As DataSet

            ' **********************************************************************************
            '   Function    :   GetDataViewFromCSV()
            '   Purpose     :   This function will upload the data from CSV file to Dataview
            '   Parameters  :   ByRef CurrFileToUpload AS String 
            '                       CurrFileToUpload = File Name
            '   Author      :   JOHN
            '   Date        :   MAY 2015 27
            ' **********************************************************************************

            Dim MyConnection As System.Data.OleDb.OleDbConnection = Nothing
            Dim MyCommand As System.Data.OleDb.OleDbDataAdapter = Nothing
            Dim workbook As String = String.Empty
            Dim sw As StreamWriter = Nothing
            Dim oDS_BudgetTemplate As New DataSet
            Dim xl As New Microsoft.Office.Interop.Excel.Application

            Dim xlsheet As Microsoft.Office.Interop.Excel.Worksheet
            Dim xlwbook As Microsoft.Office.Interop.Excel.Workbook
            '  Dim oDT As DataTable
            Dim sFuncName As String = String.Empty
            p_sBudgetType = String.Empty

            Try
                sFuncName = "GetDatasetFromExcel()"
                WriteIntoEditBox(oForm, "Starting Function - '" & sFuncName & "' ", sErrDesc)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

                MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.ACE.OLEDB.12.0; " & _
                "data source=" & CurrFileToUpload & ";Extended Properties=""Excel 12.0;HDR=Yes;IMEX=1""")

                xl.DisplayAlerts = False
                xlwbook = xl.Workbooks.Open(CurrFileToUpload, [ReadOnly]:=True, Notify:=False, IgnoreReadOnlyRecommended:=True)
                xlsheet = xlwbook.Sheets.Item(1)
                WriteIntoEditBox(oForm, "Attempting Information from the Budget Template .... ", sErrDesc)
                For Each sht In xlwbook.Worksheets
                    If sht.Visible = True Then
                        workbook = sht.Name
                        MyCommand = New System.Data.OleDb.OleDbDataAdapter("SELECT * FROM [" & workbook & "$]", MyConnection)
                        MyCommand.TableMappings.Add("Table", workbook)
                        MyCommand.Fill(oDS_BudgetTemplate)
                        WriteIntoEditBox(oForm, "Fetching Information from WorkBook - '" & workbook & "' ", sErrDesc)
                    End If
                Next

                p_sBudgetType = Left(oDS_BudgetTemplate.Tables(0).Rows(0).Item(0).ToString.Trim, 2)

                GetDatasetFromExcel = oDS_BudgetTemplate
                xl.DisplayAlerts = True
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
                WriteIntoEditBox(oForm, "Completed with SUCCESS ..... ", sErrDesc)

            Catch ex As Exception
                sErrDesc = ex.Message
                p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error occured while reading content of  " & ex.Message, sFuncName)
                Call WriteToLogFile(ex.Message, sFuncName)
                Return Nothing
            Finally

                ''   xl.ActiveWorkbook.Close(True)
                xl.Quit()
                '' xl.Workbooks.Close()
                xlwbook = Nothing
                xl = Nothing
                MyCommand.Dispose()
                MyConnection.Dispose()
            End Try

        End Function

        Public Function WriteIntoEditBox(ByRef oform As SAPbouiCOM.Form, ByVal sMsg As String, ByRef sErrDesc As String)
            Try
                If String.IsNullOrEmpty(oform.Items.Item("1000001").Specific.String.ToString.Trim()) Then
                    oform.Items.Item("1000001").Specific.String = sMsg
                Else
                    oform.Items.Item("1000001").Specific.String = oform.Items.Item("1000001").Specific.String & vbNewLine & sMsg & vbNewLine
                End If

            Catch ex As Exception
                p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End Try
        End Function

        Public Function StreamlineBudgetInformation_OU(ByRef oForm As SAPbouiCOM.Form, ByVal oDS_Budget As DataSet, ByRef sErrDesc As String) As DataTable

            ' **********************************************************************************
            '   Function    :   StreamlineBudgetInformation()
            '   Purpose     :   This function will upload the data from CSV file to Dataview
            '   Parameters  :   ByRef CurrFileToUpload AS String 
            '                       CurrFileToUpload = File Name
            '   Author      :   JOHN
            '   Date        :   MAY 2015 27
            ' **********************************************************************************

            Dim oDT_BudgetFinal As New DataTable
            Dim sFuncName As String = String.Empty
            Dim oDV As DataView = Nothing
            Dim sBudgetType As String = String.Empty
            Dim sBudgetName As String = String.Empty
            Dim sBudgetPeriod As String = String.Empty
            Dim sCostCenter As String = String.Empty

            Try
                sFuncName = "StreamlineBudgetInformation()"
                WriteIntoEditBox(oForm, "Starting Function - '" & sFuncName & "' ", sErrDesc)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

                oDT_BudgetFinal.Columns.Add("BudgetType", GetType(String))
                oDT_BudgetFinal.Columns.Add("BudgetName", GetType(String))
                oDT_BudgetFinal.Columns.Add("BudgetPeriod", GetType(String))
                oDT_BudgetFinal.Columns.Add("GLAccount", GetType(String))
                oDT_BudgetFinal.Columns.Add("CostCenter", GetType(String))
                oDT_BudgetFinal.Columns.Add("BudgetAmount", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month1", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month2", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month3", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month4", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month5", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month6", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month7", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month8", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month9", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month10", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month11", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month12", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("ProjectCode", GetType(String))
                oDT_BudgetFinal.Columns.Add("BUCode", GetType(String))
                oDT_BudgetFinal.Columns.Add("sAmount", GetType(Decimal))
               

                For imjs As Integer = 0 To oDS_Budget.Tables.Count - 1
                    If oDS_Budget.Tables(imjs).Rows.Count > 1 Then
                        oDV = New DataView(oDS_Budget.Tables(imjs))
                        sBudgetType = Left(oDV.Item(0).Row(0).ToString.Trim, 2)
                        sBudgetName = oDV.Item(3).Row("F3").ToString.Trim
                        sBudgetPeriod = oDV.Item(5).Row("F3").ToString.Trim
                        sCostCenter = oDV.Item(2).Row("F3").ToString.Trim
                        WriteIntoEditBox(oForm, "Validating Information for CostCenter " & sCostCenter, sErrDesc)
                        For imjd As Integer = 10 To oDV.Count - 1
                            If Not String.IsNullOrEmpty(oDV.Item(imjd).Row(0).ToString) Then
                                Dim dd = oDV.Item(imjd).Row("F16")
                                oDT_BudgetFinal.Rows.Add(sBudgetType, sBudgetName, sBudgetPeriod, oDV.Item(imjd).Row(0), sCostCenter, oDV.Item(imjd).Row("F16"), oDV.Item(imjd).Row("F4"), oDV.Item(imjd).Row("F5"), _
                                                         oDV.Item(imjd).Row("F6"), oDV.Item(imjd).Row("F7"), oDV.Item(imjd).Row("F8"), oDV.Item(imjd).Row("F9"), oDV.Item(imjd).Row("F10"), _
                                                         oDV.Item(imjd).Row("F11"), oDV.Item(imjd).Row("F12"), oDV.Item(imjd).Row("F13"), oDV.Item(imjd).Row("F14"), oDV.Item(imjd).Row("F15"), "", "", 0)

                            End If
                        Next imjd
                    End If

                Next imjs

                StreamlineBudgetInformation_OU = SetDefaultForNull(oDT_BudgetFinal)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
                WriteIntoEditBox(oForm, "Completed with SUCCESS ..... ", sErrDesc)

            Catch ex As Exception
                sErrDesc = ex.Message
                p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error occured while reading content of  " & ex.Message, sFuncName)
                Call WriteToLogFile(ex.Message, sFuncName)
                Return Nothing
            Finally

            End Try

        End Function

        Public Function StreamlineBudgetInformation_BU(ByRef oForm As SAPbouiCOM.Form, ByVal oDS_Budget As DataSet, ByRef sErrDesc As String) As DataTable

            ' **********************************************************************************
            '   Function    :   StreamlineBudgetInformation_BU()
            '   Purpose     :   This function will upload the data from CSV file to Dataview
            '   Parameters  :   ByRef CurrFileToUpload AS String 
            '                       CurrFileToUpload = File Name
            '   Author      :   JOHN
            '   Date        :   MAY 2015 27
            ' **********************************************************************************

            Dim oDT_BudgetFinal As New DataTable
            Dim sFuncName As String = String.Empty
            Dim oDV As DataView = Nothing
            Dim sBudgetType As String = String.Empty
            Dim sBudgetName As String = String.Empty
            Dim sBudgetPeriod As String = String.Empty
            Dim sCostCenter As String = String.Empty
            Dim sBUCode As String = String.Empty
            Dim dSpltAmount As Double = 0.0

            Try
                sFuncName = "StreamlineBudgetInformation_BU()"
                WriteIntoEditBox(oForm, "Starting Function - '" & sFuncName & "' ", sErrDesc)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

                oDT_BudgetFinal.Columns.Add("BudgetType", GetType(String))
                oDT_BudgetFinal.Columns.Add("BudgetName", GetType(String))
                oDT_BudgetFinal.Columns.Add("GLAccount", GetType(String))
                oDT_BudgetFinal.Columns.Add("BudgetPeriod", GetType(String))
                oDT_BudgetFinal.Columns.Add("BUCode", GetType(String))
                oDT_BudgetFinal.Columns.Add("BudgetAmount", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("CostCenter", GetType(String))
                oDT_BudgetFinal.Columns.Add("sAmount", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month1", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month2", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month3", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month4", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month5", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month6", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month7", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month8", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month9", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month10", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month11", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month12", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("ProjectCode", GetType(String))

                For imjs As Integer = 0 To oDS_Budget.Tables.Count - 1
                    If oDS_Budget.Tables(imjs).Rows.Count > 1 Then
                        oDV = New DataView(oDS_Budget.Tables(imjs))
                        sBudgetType = Left(oDV.Item(0).Row(0).ToString.Trim, 2)
                        sBudgetName = oDV.Item(3).Row("F3").ToString.Trim
                        sBudgetPeriod = oDV.Item(5).Row("F3").ToString.Trim
                        sBUCode = oDV.Item(2).Row("F3").ToString.Trim
                        WriteIntoEditBox(oForm, "Validating Information for BU Code " & sBUCode, sErrDesc)
                        For imjd As Integer = 10 To oDV.Count - 1
                            If Not String.IsNullOrEmpty(oDV.Item(imjd).Row(0).ToString) Then

                                ''oDT_BudgetFinal.Rows.Add(sBudgetType, sBudgetName, oDV.Item(imjd).Row(0), sBudgetPeriod, sBUCode, oDV.Item(imjd).Row("F18"), oDV.Item(imjd).Row("F4"), oDV.Item(imjd).Row("F5"), _
                                ''                         oDV.Item(imjd).Row("F6").ToString, oDV.Item(imjd).Row("F7"), oDV.Item(imjd).Row("F8"), oDV.Item(imjd).Row("F9"), oDV.Item(imjd).Row("F10"), _
                                ''                         oDV.Item(imjd).Row("F11"), oDV.Item(imjd).Row("F12"), oDV.Item(imjd).Row("F13"), oDV.Item(imjd).Row("F14"), oDV.Item(imjd).Row("F15"), oDV.Item(imjd).Row("F16"), oDV.Item(imjd).Row("F17"), "")

                                If Not String.IsNullOrEmpty(oDV.Item(imjd).Row("F5").ToString) Then
                                    dSpltAmount = FormatNumber(oDV.Item(imjd).Row("F5") / 12, 2)
                                Else
                                    dSpltAmount = 0
                                End If

                                oDT_BudgetFinal.Rows.Add(sBudgetType, sBudgetName, oDV.Item(imjd).Row(0), sBudgetPeriod, sBUCode, oDV.Item(imjd).Row("F5"), oDV.Item(imjd).Row("F4"), oDV.Item(imjd).Row("F5"), _
                                                        dSpltAmount, dSpltAmount, dSpltAmount, dSpltAmount, dSpltAmount, dSpltAmount, dSpltAmount, dSpltAmount, dSpltAmount, dSpltAmount, dSpltAmount, dSpltAmount, "")
                            End If
                        Next imjd
                    End If

                Next imjs

                StreamlineBudgetInformation_BU = SetDefaultForNull(oDT_BudgetFinal)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
                WriteIntoEditBox(oForm, "Completed with SUCCESS ..... ", sErrDesc)

            Catch ex As Exception
                sErrDesc = ex.Message
                p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error occured while reading content of  " & ex.Message, sFuncName)
                Call WriteToLogFile(ex.Message, sFuncName)
                Return Nothing
            Finally

            End Try

        End Function

        Public Function StreamlineBudgetInformation_PROLD(ByRef oForm As SAPbouiCOM.Form, ByVal oDS_Budget As DataSet, ByRef sErrDesc As String) As DataTable

            ' **********************************************************************************
            '   Function    :   StreamlineBudgetInformation_PR()
            '   Purpose     :   This function will upload the data from CSV file to Dataview
            '   Parameters  :   ByRef CurrFileToUpload AS String 
            '                       CurrFileToUpload = File Name
            '   Author      :   JOHN
            '   Date        :   MAY 2015 27
            ' **********************************************************************************

            Dim oDT_BudgetFinal As New DataTable
            Dim sFuncName As String = String.Empty
            Dim oDV As DataView = Nothing
            Dim sBudgetType As String = String.Empty
            Dim sBudgetName As String = String.Empty
            Dim sBudgetPeriod As String = String.Empty
            Dim sBUCode As String = String.Empty
            Dim sCostCenter As String = String.Empty
            Dim sPrjCode As String = String.Empty
            Dim iIndexRange As Integer = 4
            Dim iIndexCount As Integer = 0
            Dim scolumnName As String = String.Empty

            Try
                sFuncName = "StreamlineBudgetInformation_PROLD()"
                WriteIntoEditBox(oForm, "Starting Function - '" & sFuncName & "' ", sErrDesc)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

                oDT_BudgetFinal.Columns.Add("BudgetType", GetType(String))
                oDT_BudgetFinal.Columns.Add("BudgetName", GetType(String))
                oDT_BudgetFinal.Columns.Add("GLAccount", GetType(String))
                oDT_BudgetFinal.Columns.Add("BudgetPeriod", GetType(String))
                oDT_BudgetFinal.Columns.Add("BUCode", GetType(String))
                oDT_BudgetFinal.Columns.Add("BudgetAmount", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("CostCenter", GetType(String))
                oDT_BudgetFinal.Columns.Add("sAmount", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month1", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month2", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month3", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month4", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month5", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month6", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month7", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month8", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month9", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month10", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month11", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month12", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("ProjectCode", GetType(String))


                For imjs As Integer = 0 To oDS_Budget.Tables.Count - 1
                    If oDS_Budget.Tables(imjs).Rows.Count > 1 Then
                        oDV = New DataView(oDS_Budget.Tables(imjs))
                        For imjd As Integer = iIndexRange To 15
                            scolumnName = "F" & Trim(Str(imjd))
                            If oDV.Item(9).Row(scolumnName.Trim()).ToString.Trim <> "Total" Then
                                iIndexCount += 1
                            Else
                                Exit For
                            End If
                        Next imjd

                        sBudgetType = Left(oDV.Item(0).Row(0).ToString.Trim, 2)
                        sBudgetName = oDV.Item(3).Row("F3").ToString.Trim
                        sBudgetPeriod = oDV.Item(5).Row("F3").ToString.Trim
                        sBUCode = oDV.Item(2).Row("F3").ToString.Trim

                        For icount As Integer = iIndexRange To iIndexCount + iIndexRange - 1
                            sPrjCode = oDV.Item(9).Row("F" & Trim(Str(icount))).ToString.Trim
                            WriteIntoEditBox(oForm, "Validating Information for Project Code " & sPrjCode, sErrDesc)

                            For imjd As Integer = 10 To oDV.Count - 1
                                If Not String.IsNullOrEmpty(oDV.Item(imjd).Row(0).ToString) Then

                                    ''oDT_BudgetFinal.Rows.Add(sBudgetType, sBudgetName, oDV.Item(imjd).Row(0), sBudgetPeriod, sBUCode, oDV.Item(imjd).Row("F18"), oDV.Item(imjd).Row("F4"), oDV.Item(imjd).Row("F5"), _
                                    ''                         oDV.Item(imjd).Row("F6").ToString, oDV.Item(imjd).Row("F7"), oDV.Item(imjd).Row("F8"), oDV.Item(imjd).Row("F9"), oDV.Item(imjd).Row("F10"), _
                                    ''                         oDV.Item(imjd).Row("F11"), oDV.Item(imjd).Row("F12"), oDV.Item(imjd).Row("F13"), oDV.Item(imjd).Row("F14"), oDV.Item(imjd).Row("F15"), oDV.Item(imjd).Row("F16"), oDV.Item(imjd).Row("F17"), "")
                                    If icount <> iIndexCount + iIndexRange - 1 Then
                                        oDT_BudgetFinal.Rows.Add(sBudgetType, sBudgetName, oDV.Item(imjd).Row(0), sBudgetPeriod, "", oDV.Item(imjd).Row("F" & Trim(Str(icount))), "", 0, _
                                                          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, sPrjCode)
                                    Else
                                        oDT_BudgetFinal.Rows.Add(sBudgetType, sBudgetName, oDV.Item(imjd).Row(0), sBudgetPeriod, "", oDV.Item(imjd).Row("F" & Trim(Str(icount))), "", 0, _
                                                          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "")
                                    End If

                                End If
                            Next imjd
                        Next

                    End If

                Next imjs

                StreamlineBudgetInformation_PROLD = SetDefaultForNull(oDT_BudgetFinal)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
                WriteIntoEditBox(oForm, "Completed with SUCCESS ..... ", sErrDesc)

            Catch ex As Exception
                sErrDesc = ex.Message
                p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error occured while reading content of  " & ex.Message, sFuncName)
                Call WriteToLogFile(ex.Message, sFuncName)
                Return Nothing
            Finally

            End Try

        End Function


        Public Function StreamlineBudgetInformation_PR(ByRef oForm As SAPbouiCOM.Form, ByVal oDS_Budget As DataSet, ByRef sErrDesc As String) As DataTable

            ' **********************************************************************************
            '   Function    :   StreamlineBudgetInformation_PR()
            '   Purpose     :   This function will upload the data from CSV file to Dataview
            '   Parameters  :   ByRef CurrFileToUpload AS String 
            '                       CurrFileToUpload = File Name
            '   Author      :   JOHN
            '   Date        :   MAY 2015 27
            ' **********************************************************************************

            Dim oDT_BudgetFinal As New DataTable
            Dim sFuncName As String = String.Empty
            Dim oDV As DataView = Nothing
            Dim sBudgetType As String = String.Empty
            Dim sBudgetAmount As Double = 0.0
            Dim sBudgetName As String = String.Empty
            Dim sBudgetPeriod As String = String.Empty
            Dim sBUCode As String = String.Empty
            Dim sCostCenter As String = String.Empty
            Dim sPrjCode As String = String.Empty
            Dim iIndexRange As Integer = 4
            Dim iIndexCount As Integer = 0
            Dim scolumnName As String = String.Empty
            Dim irowcount As Integer = 0
            Dim oRow() As Data.DataRow = Nothing


            Try
                sFuncName = "StreamlineBudgetInformation_PR()"
                WriteIntoEditBox(oForm, "Starting Function - '" & sFuncName & "' ", sErrDesc)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

                oDT_BudgetFinal.Columns.Add("BudgetType", GetType(String))
                oDT_BudgetFinal.Columns.Add("BudgetName", GetType(String))
                oDT_BudgetFinal.Columns.Add("GLAccount", GetType(String))
                oDT_BudgetFinal.Columns.Add("BudgetPeriod", GetType(String))
                oDT_BudgetFinal.Columns.Add("BUCode", GetType(String))
                oDT_BudgetFinal.Columns.Add("BudgetAmount", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("CostCenter", GetType(String))
                oDT_BudgetFinal.Columns.Add("sAmount", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month1", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month2", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month3", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month4", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month5", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month6", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month7", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month8", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month9", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month10", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month11", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("Month12", GetType(Decimal))
                oDT_BudgetFinal.Columns.Add("ProjectCode", GetType(String))


                For imjs As Integer = 0 To oDS_Budget.Tables.Count - 1
                    iIndexCount = 0
                    If oDS_Budget.Tables(imjs).Rows.Count > 1 Then
                        oDV = New DataView(oDS_Budget.Tables(imjs))
                        For imjd As Integer = iIndexRange To 100
                            scolumnName = "F" & Trim(Str(imjd))
                            If oDV.Item(9).Row(scolumnName.Trim()).ToString.Trim <> "Total" Then
                                iIndexCount += 1
                            Else
                                Exit For
                            End If
                        Next imjd

                        sBudgetType = Left(oDV.Item(0).Row(0).ToString.Trim, 2)
                        sBudgetName = oDV.Item(3).Row("F3").ToString.Trim
                        sBudgetPeriod = oDV.Item(5).Row("F3").ToString.Trim
                        sBUCode = oDV.Item(2).Row("F3").ToString.Trim
                        irowcount = iIndexCount + iIndexRange - 1

                        For icount As Integer = iIndexRange To irowcount
                            sPrjCode = oDV.Item(9).Row("F" & Trim(Str(icount))).ToString.Trim
                            WriteIntoEditBox(oForm, "Validating Information for Project Code " & sPrjCode, sErrDesc)

                            For imjd As Integer = 10 To oDV.Count - 1
                                If Not String.IsNullOrEmpty(oDV.Item(imjd).Row(0).ToString) Then

                                    If String.IsNullOrEmpty(oDV.Item(imjd).Row("F" & Trim(Str(icount))).ToString) Then
                                        sBudgetAmount = 0
                                    Else
                                        sBudgetAmount = oDV.Item(imjd).Row("F" & Trim(Str(icount))).ToString
                                    End If

                                    ''oDT_BudgetFinal.Rows.Add(sBudgetType, sBudgetName, oDV.Item(imjd).Row(0), sBudgetPeriod, sBUCode, oDV.Item(imjd).Row("F18"), oDV.Item(imjd).Row("F4"), oDV.Item(imjd).Row("F5"), _
                                    ''                         oDV.Item(imjd).Row("F6").ToString, oDV.Item(imjd).Row("F7"), oDV.Item(imjd).Row("F8"), oDV.Item(imjd).Row("F9"), oDV.Item(imjd).Row("F10"), _
                                    ''                         oDV.Item(imjd).Row("F11"), oDV.Item(imjd).Row("F12"), oDV.Item(imjd).Row("F13"), oDV.Item(imjd).Row("F14"), oDV.Item(imjd).Row("F15"), oDV.Item(imjd).Row("F16"), oDV.Item(imjd).Row("F17"), "")
                                    If UCase(Left(sPrjCode, 3)) = "NON" Then

                                        oRow = oDT_BudgetFinal.Select("GLAccount = '" & oDV.Item(imjd).Row(0).ToString.Trim() & "' and BUCode = '" & sBUCode & "'")
                                        If oRow.Count > 0 Then
                                            oRow(0)("BudgetAmount") += sBudgetAmount
                                        Else
                                            oDT_BudgetFinal.Rows.Add(sBudgetType, sBudgetName, oDV.Item(imjd).Row(0), sBudgetPeriod, sBUCode, sBudgetAmount, "", 0, _
                                                          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "")
                                        End If
                                    Else

                                        oRow = oDT_BudgetFinal.Select("GLAccount = '" & oDV.Item(imjd).Row(0).ToString.Trim() & "' and ProjectCode = '" & sPrjCode & "'")
                                        If oRow.Count > 0 Then
                                            oRow(0)("BudgetAmount") += sBudgetAmount
                                        Else
                                            oDT_BudgetFinal.Rows.Add(sBudgetType, sBudgetName, oDV.Item(imjd).Row(0), sBudgetPeriod, "", sBudgetAmount, "", 0, _
                                                                                                     0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, sPrjCode)
                                        End If
                                    End If
                                End If
                            Next imjd
                        Next

                    End If

                Next imjs

                StreamlineBudgetInformation_PR = SetDefaultForNull(oDT_BudgetFinal)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
                WriteIntoEditBox(oForm, "Completed with SUCCESS ..... ", sErrDesc)

            Catch ex As Exception
                sErrDesc = ex.Message
                p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error occured while reading content of  " & ex.Message, sFuncName)
                Call WriteToLogFile(ex.Message, sFuncName)
                Return Nothing
            Finally

            End Try

        End Function

        Public Function SetDefaultForNull(ByVal dataTable As DataTable) As DataTable
            For Each row As DataRow In dataTable.Rows
                For Each col As DataColumn In dataTable.Columns
                    Dim value As Object = row(col)
                    If IsDBNull(value) Then
                        Dim dataType As String = col.DataType.ToString

                        Select Case dataType
                            Case "System.DateTime"
                                value = New DateTime
                            Case "System.Decimal", "System.Int16", "System.Int32", "System.Int64"
                                value = 0
                            Case "System.String"
                                value = String.Empty
                            Case "System.Boolean"
                                value = False
                            Case Else
                                value = 0
                        End Select
                        row(col) = value
                    End If
                Next
            Next
            Return dataTable
        End Function

        Function Loading_AgingDetails(ByRef FormUID As SAPbouiCOM.Form, ByRef oApplication As SAPbouiCOM.Application _
                                   , ByRef oCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long

            Dim sFuncName As String = String.Empty
            Dim oMatrix As SAPbouiCOM.Matrix = Nothing
            Dim oRset As SAPbobsCOM.Recordset = Nothing
            Dim sQry As String = String.Empty
            Dim AgingDate As String = String.Empty


            Try
                sFuncName = "Loading_AgingDetails()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AgingDate()", sFuncName)
                '' AgingDate = GateDate(FormUID.Items.Item("Item_7").Specific.String, oCompany)
                ''sQry = "SELECT T0.[U_AE_BPCode], T0.[U_AE_BPName],T0.[U_AE_Balance], T0.[U_AE_Date], T0.[U_AE_Email] [Free_Text], T0.[U_AE_Status] [CardFName], T0.[U_AE_ErrMsg] FROM [dbo].[@AE_AGINGLOG]  T0 WHERE T0.[U_AE_Date]  = '" & AgingDate & "'"
                ''oRset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                ''If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Query for Aging Table " & sQry, sFuncName)
                ''oRset.DoQuery(sQry)

                sQry = "AB_SOA_OS_SP004'" & FormUID.Items.Item("BPFrom").Specific.String & "','" & FormUID.Items.Item("BPTo").Specific.String & "','" & AgingDate & "','1'"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("SQL " & sQry, sFuncName)


                oMatrix = FormUID.Items.Item("Item_8").Specific

                Try
                    FormUID.DataSources.DataTables.Add("OCRD")
                Catch ex As Exception

                End Try

                FormUID.DataSources.DataTables.Item("OCRD").ExecuteQuery(sQry)
                oMatrix.Clear()
                FormUID.Items.Item("Item_8").Specific.columns.item("Col_1").databind.bind("OCRD", "CardCode")
                FormUID.Items.Item("Item_8").Specific.columns.item("Col_2").databind.bind("OCRD", "CardName")
                FormUID.Items.Item("Item_8").Specific.columns.item("Col_3").databind.bind("OCRD", "Balance")
                FormUID.Items.Item("Item_8").Specific.columns.item("Col_4").databind.bind("OCRD", "Free_Text")
                FormUID.Items.Item("Item_8").Specific.LoadFromDataSource()


                Loading_AgingDetails = RTN_SUCCESS
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

            Catch ex As Exception
                p_oSBOApplication.StatusBar.SetText("HeadValidation Function : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                sErrDesc = ex.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                Loading_AgingDetails = RTN_ERROR
            End Try
            Return RTN_SUCCESS
        End Function

        Public Function GetSingleValue(ByVal sAccountCode As String, ByVal sGDC As String) As String
            Try
                Dim objRS As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim sSqlString As String = String.Empty

                If sGDC = "G" Then
                    sSqlString = "SELECT T0.U_BibbySGCode [Name] FROM [dbo].[@BIBBY_ACCT_MAPPING]  T0 WHERE T0.U_BibbyAFCode ='" & sAccountCode & "'"
                Else
                    sSqlString = "SELECT T0.U_BibbySGCode [Name] FROM [dbo].[@BIBBY_ACCT_MAPPING]  T0 WHERE T0.U_BibbyAFCode ='" & sGDC & "'"
                End If

                objRS.DoQuery(sSqlString)
                If objRS.RecordCount > 0 Then
                    Return objRS.Fields.Item(0).Value.ToString
                End If
            Catch ex As Exception
                Return ""
            End Try
            Return Nothing
        End Function

        Public Function Del_schema(ByVal csvFileFolder As String) As Long

            ' ***********************************************************************************
            '   Function   :    Del_schema()
            '   Purpose    :    This function is handles - Delete the Schema file
            '   Parameters :    ByVal csvFileFolder As String
            '                       csvFileFolder = Passing file name
            '   Author     :    JOHN
            '   Date       :    26/06/2014 
            '   Change     :   
            '                   
            ' ***********************************************************************************
            Dim sFuncName As String = String.Empty
            Try
                sFuncName = "Del_schema()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function...", sFuncName)
                Console.WriteLine("Starting Function... " & sFuncName)

                Dim FileToDelete As String
                FileToDelete = csvFileFolder & "\\schema.ini"
                If System.IO.File.Exists(FileToDelete) = True Then
                    System.IO.File.Delete(FileToDelete)
                End If
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
                Console.WriteLine("Completed with SUCCESS " & sFuncName)
                Del_schema = RTN_SUCCESS
            Catch ex As Exception
                WriteToLogFile(ex.Message, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                Console.WriteLine("Completed with Error " & sFuncName)
                Del_schema = RTN_ERROR
            End Try
        End Function

        Public Function Create_schema(ByVal csvFileFolder As String, ByVal FileName As String) As Long

            ' ***********************************************************************************
            '   Function   :    Create_schema()
            '   Purpose    :    This function is handles - Create the Schema file
            '   Parameters :    ByVal csvFileFolder As String
            '                       csvFileFolder = Passing file name
            '   Author     :    JOHN
            '   Date       :    26/06/2014 
            '   Change     :   
            '                   
            ' ***********************************************************************************
            Dim sFuncName As String = String.Empty
            Try
                sFuncName = "Create_schema()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function...", sFuncName)
                Console.WriteLine("Starting Function... " & sFuncName)

                Dim csvFileName As String = FileName
                Dim fsOutput As FileStream = New FileStream(csvFileFolder & "\\schema.ini", FileMode.Create, FileAccess.Write)
                Dim srOutput As StreamWriter = New StreamWriter(fsOutput)
                Dim s1, s2, s3, s4, s5 As String
                s1 = "[" & csvFileName & "]"
                s2 = "ColNameHeader=False"
                s3 = "Format=CSVDelimited"
                s4 = "MaxScanRows=0"
                s5 = "CharacterSet=OEM"
                srOutput.WriteLine(s1.ToString() + ControlChars.Lf + s2.ToString() + ControlChars.Lf + s3.ToString() + ControlChars.Lf + s4.ToString() + ControlChars.Lf)
                srOutput.Close()
                fsOutput.Close()

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
                Console.WriteLine("Completed with SUCCESS " & sFuncName)
                Create_schema = RTN_SUCCESS

            Catch ex As Exception
                WriteToLogFile(ex.Message, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                Console.WriteLine("Completed with Error " & sFuncName)
                Create_schema = RTN_ERROR
            End Try

        End Function

        Public Function GetDate(ByVal sDate As String, ByRef oCompany As SAPbobsCOM.Company) As String

            Dim dateValue As DateTime
            Dim DateString As String = String.Empty
            Dim sSQL As String = String.Empty
            Dim oRs As SAPbobsCOM.Recordset
            Dim sDatesep As String

            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            sSQL = "SELECT DateFormat,DateSep FROM OADM"

            oRs.DoQuery(sSQL)

            If Not oRs.EoF Then
                sDatesep = oRs.Fields.Item("DateSep").Value

                Select Case oRs.Fields.Item("DateFormat").Value
                    Case 0
                        If Date.TryParseExact(sDate, "dd" & sDatesep & "MM" & sDatesep & "yy", _
                           New CultureInfo("en-US"), _
                           DateTimeStyles.None, _
                           dateValue) Then
                            DateString = dateValue.ToString("yyyyMMdd")
                        End If
                    Case 1
                        If Date.TryParseExact(sDate, "dd" & sDatesep & "MM" & sDatesep & "yyyy", _
                           New CultureInfo("en-US"), _
                           DateTimeStyles.None, _
                           dateValue) Then
                            DateString = dateValue.ToString("yyyyMMdd")
                        End If
                    Case 2
                        If Date.TryParseExact(sDate, "MM" & sDatesep & "dd" & sDatesep & "yy", _
                            New CultureInfo("en-US"), _
                            DateTimeStyles.None, _
                            dateValue) Then
                            DateString = dateValue.ToString("yyyyMMdd")
                        End If
                    Case 3
                        If Date.TryParseExact(sDate, "MM" & sDatesep & "dd" & sDatesep & "yyyy", _
                            New CultureInfo("en-US"), _
                            DateTimeStyles.None, _
                            dateValue) Then
                            DateString = dateValue.ToString("yyyyMMdd")
                        End If
                    Case 4
                        If Date.TryParseExact(sDate, "yyyy" & sDatesep & "MM" & sDatesep & "dd", _
                            New CultureInfo("en-US"), _
                            DateTimeStyles.None, _
                            dateValue) Then
                            DateString = dateValue.ToString("yyyyMMdd")
                        End If
                    Case 5
                        If Date.TryParseExact(sDate, "dd" & sDatesep & "MMMM" & sDatesep & "yyyy", _
                            New CultureInfo("en-US"), _
                            DateTimeStyles.None, _
                            dateValue) Then
                            DateString = dateValue.ToString("yyyyMMdd")
                        End If
                    Case 6
                        If Date.TryParseExact(sDate, "yy" & sDatesep & "MM" & sDatesep & "dd", _
                            New CultureInfo("en-US"), _
                            DateTimeStyles.None, _
                            dateValue) Then
                            DateString = dateValue.ToString("yyyyMMdd")
                        End If
                    Case Else
                        DateString = dateValue.ToString("yyyyMMdd")
                End Select

            End If

            Return DateString

        End Function

        Public Function PostDate(ByRef oCompany As SAPbobsCOM.Company) As String

            Dim DateString As String = String.Empty
            Dim sSQL As String = String.Empty
            Dim oRs As SAPbobsCOM.Recordset
            Dim sDatesep As String

            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            sSQL = "SELECT DateFormat,DateSep FROM OADM"

            oRs.DoQuery(sSQL)

            If Not oRs.EoF Then
                sDatesep = oRs.Fields.Item("DateSep").Value

                Select Case oRs.Fields.Item("DateFormat").Value
                    Case 0

                        DateString = Format(Now.Date, "dd" & sDatesep & "MM" & sDatesep & "yy")

                    Case 1
                        DateString = Format(Now.Date, "dd" & sDatesep & "MM" & sDatesep & "yyyy")

                    Case 2
                        DateString = Format(Now.Date, "MM" & sDatesep & "dd" & sDatesep & "yy")
                    Case 3
                        DateString = Format(Now.Date, "MM" & sDatesep & "dd" & sDatesep & "yyyy")
                    Case 4
                        DateString = Format(Now.Date, "yyyy" & sDatesep & "MM" & sDatesep & "dd")
                    Case 5
                        DateString = Format(Now.Date, "dd" & sDatesep & "MMMM" & sDatesep & "yyyy")
                    Case 6
                        DateString = Format(Now.Date, "yy" & sDatesep & "MM" & sDatesep & "dd")
                End Select

            End If

            Return DateString

        End Function

        Public Function Write_TextFile(ByVal oDT_FinalResult As DataTable, ByVal sPAth As String, ByVal scheck As String, ByRef sErrDesc As String) As Long
            Try
                Dim sFuncName As String = String.Empty
                Dim irow As Integer
                Dim sFileName As String = "\SAPGeneratedFile.txt"
                Dim sbuffer As String = String.Empty

                sFuncName = "Write_TextFile()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

                If File.Exists(sPAth & sFileName) Then
                    Try
                        File.Delete(sPAth & sFileName)
                    Catch ex As Exception
                    End Try
                End If

                Dim sw As StreamWriter = New StreamWriter(sPAth & sFileName)
                ' Add some text to the file.

                For imjs = 0 To oDT_FinalResult.Rows.Count - 1

                    sw.WriteLine(oDT_FinalResult.Rows(imjs).Item("AcctCode").ToString & vbTab & oDT_FinalResult.Rows(imjs).Item("RefDate").ToString & vbTab & oDT_FinalResult.Rows(imjs).Item("OU Code").ToString _
                                 & oDT_FinalResult.Rows(imjs).Item("Entity").ToString & vbTab & oDT_FinalResult.Rows(imjs).Item("DC").ToString & vbTab & oDT_FinalResult.Rows(imjs).Item("Amount").ToString)

                Next imjs
                sw.Close()
                If scheck = "Y" Then
                    Process.Start(sPAth & sFileName)
                End If
                '' Process.Start(sPath & sFileName)

                Write_TextFile = RTN_SUCCESS
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With SUCCESS ", sFuncName)

            Catch ex As Exception
                Write_TextFile = RTN_ERROR
                sErrDesc = ex.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            End Try

        End Function

        Public Function ConvertStringToDate(ByRef sDate As String) As Date
            Try
                'Dim iIndex As Integer = 0
                'Dim iDay As String
                'Dim iMonth As String
                Dim sMonth() As String

                sMonth = sDate.Split("/")
                Return sMonth(2) & "/" & sMonth(1).PadLeft(2, "0"c) & "/" & sMonth(0).PadLeft(2, "0"c)
            Catch ex As Exception
                p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Return "1/1/1"
            End Try

        End Function

        Public Function ConvertRecordset(ByVal SAPRecordset As SAPbobsCOM.Recordset, ByRef sErrDesc As String)

            '\ This function will take an SAP recordset from the SAPbobsCOM library and convert it to a more
            '\ easily used ADO.NET datatable which can be used for data binding much easier.


            Dim NewCol As DataColumn
            Dim NewRow As DataRow
            Dim ColCount As Integer

            Try
                For ColCount = 0 To SAPRecordset.Fields.Count - 1
                    NewCol = New DataColumn(SAPRecordset.Fields.Item(ColCount).Name)
                    dtTable.Columns.Add(NewCol)
                Next

                Do Until SAPRecordset.EoF

                    NewRow = dtTable.NewRow
                    'populate each column in the row we're creating
                    For ColCount = 0 To SAPRecordset.Fields.Count - 1

                        NewRow.Item(SAPRecordset.Fields.Item(ColCount).Name) = SAPRecordset.Fields.Item(ColCount).Value

                    Next

                    'Add the row to the datatable
                    dtTable.Rows.Add(NewRow)


                    SAPRecordset.MoveNext()
                Loop

                Return dtTable

            Catch ex As Exception
                sErrDesc = ex.Message
                p_oSBOApplication.StatusBar.SetText(ex.ToString & Chr(10) & "Error converting SAP Recordset to DataTable", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Function
            End Try


        End Function

#Region "   GST Audit File Common Functions    "

        Function HeaderValidation_AuditFile(ByVal FormUID As SAPbouiCOM.Form, ByRef sErrDesc As String) As Long
            Dim sFuncName As String = String.Empty
            Try
                sFuncName = "HeaderValidation_AuditFile()"

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

                If FormUID.Items.Item("txtFrmDate").Specific.string = String.Empty Then
                    p_oSBOApplication.StatusBar.SetText("From Date is Missing", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    FormUID.ActiveItem = "txtFrmDate"
                    Return RTN_ERROR

                ElseIf FormUID.Items.Item("txtToDate").Specific.string = String.Empty Then
                    p_oSBOApplication.StatusBar.SetText("ToDate is Missing ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    FormUID.ActiveItem = "txtToDate"
                    Return RTN_ERROR

                    ''ElseIf FormUID.Items.Item("txtToDate").Specific.string < FormUID.Items.Item("txtFrmDate").Specific.string Then
                    ''    p_oSBOApplication.StatusBar.SetText("To Date Should not Greater than From Date ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    ''    FormUID.ActiveItem = "txtToDate"
                    ''    Return RTN_ERROR

                ElseIf FormUID.Items.Item("8").Specific.string < FormUID.Items.Item("txtToDate").Specific.string Then
                    p_oSBOApplication.StatusBar.SetText("Output File Path is Missing ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    FormUID.ActiveItem = "8"
                    Return RTN_ERROR

                End If
                HeaderValidation_AuditFile = RTN_SUCCESS
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

            Catch ex As Exception
                p_oSBOApplication.StatusBar.SetText("HeadValidation Function : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                sErrDesc = ex.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                HeaderValidation_AuditFile = RTN_ERROR
            End Try
            Return RTN_SUCCESS
        End Function


#End Region

        Public Function ConnectToTargetCompany(ByRef oCompany As SAPbobsCOM.Company, _
                                        ByVal sEntity As String, _
                                        ByVal sUsername As String, _
                                        ByVal sPassword As String, _
                                        ByRef sErrDesc As String) As Long

            ' **********************************************************************************
            '   Function    :   ConnectToTargetCompany()
            '   Purpose     :   This function will be providing to proceed the connectivity of 
            '                   using SAP DIAPI function
            '               
            '   Parameters  :   ByRef oCompany As SAPbobsCOM.Company
            '                       oCompany =  set the SAP DI Company Object
            '                   ByRef sErrDesc AS String 
            '                       sErrDesc = Error Description to be returned to calling function
            '               
            '   Return      :   0 - FAILURE
            '                   1 - SUCCESS
            '   Author      :   JOHN
            '   Date        :   MAY 2013 21
            ' **********************************************************************************

            Dim sFuncName As String = String.Empty
            Dim iRetValue As Integer = -1
            Dim iErrCode As Integer = -1
            Dim sSQL As String = String.Empty
            Dim oDs As New DataSet

            Try
                sFuncName = "ConnectToTargetCompany()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)


                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Initializing the Company Object", sFuncName)

                oCompany = New SAPbobsCOM.Company

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Assigning the representing database name", sFuncName)
                oCompany.Server = p_oDICompany.Server
                oCompany.DbServerType = p_oDICompany.DbServerType
                oCompany.LicenseServer = p_oDICompany.LicenseServer
                oCompany.CompanyDB = sEntity
                oCompany.UserName = sUsername
                oCompany.Password = sPassword
                oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English
                oCompany.UseTrusted = False
                ' oCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Connecting to the Company Database. " & sEntity, sFuncName)
                iRetValue = oCompany.Connect()
                If iRetValue <> 0 Then
                    oCompany.GetLastError(iErrCode, sErrDesc)

                    sErrDesc = String.Format("Connection to Database ({0}) {1} {2} {3}", _
                        oCompany.CompanyDB, System.Environment.NewLine, _
                                    vbTab, sErrDesc)

                    Throw New ArgumentException(sErrDesc)
                End If


                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
                Console.WriteLine("Completed with SUCCESS ", sFuncName)
                ConnectToTargetCompany = RTN_SUCCESS
            Catch ex As Exception
                sErrDesc = ex.Message
                Call WriteToLogFile(ex.Message, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                Console.WriteLine("Completed with ERROR ", sFuncName)
                ConnectToTargetCompany = RTN_ERROR
            End Try
        End Function




    End Module
End Namespace



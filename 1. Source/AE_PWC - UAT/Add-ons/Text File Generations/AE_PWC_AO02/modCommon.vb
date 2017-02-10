Option Explicit On
Imports System.Xml
Imports System.IO
Imports System.Windows.Forms
Imports System.Globalization
Imports System.Net.Mail
Imports System.Configuration



Namespace AE_PWC_AO02

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


            Try
                sFuncName = "EntityLoad()"
                Dim oMatrix As SAPbouiCOM.Matrix = oform.Items.Item("matGtTxtFi").Specific
                Dim sSQL As String = String.Empty

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)
                sSQL = "SELECT T0.[U_AB_COMCODE], T0.[U_AB_COMPANYNAME], 'Y' [Choose] FROM [dbo].[@AB_COMPANYDATA]  T0"

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("SQL " & sSQL, sFuncName)
                Try
                    oform.DataSources.DataTables.Add("U_AB_COMPANYNAME")
                Catch ex As Exception
                End Try
                oform.DataSources.DataTables.Item("U_AB_COMPANYNAME").ExecuteQuery(sSQL)
                oMatrix.Clear()
                oMatrix.Columns.Item("V_1").DataBind.Bind("U_AB_COMPANYNAME", "Choose")
                oMatrix.Columns.Item("Col_0").DataBind.Bind("U_AB_COMPANYNAME", "U_AB_COMCODE")
                oMatrix.Columns.Item("V_0").DataBind.Bind("U_AB_COMPANYNAME", "U_AB_COMPANYNAME")
                oMatrix.LoadFromDataSource()

                For imjs As Integer = 1 To oMatrix.RowCount
                    oMatrix.Columns.Item("V_-1").Cells.Item(imjs).Specific.String = imjs
                Next imjs

                oMatrix.Columns.Item("V_0").Editable = False
                oMatrix.Columns.Item("Col_0").Editable = False
                oMatrix.AutoResizeColumns()

                EntityLoad = RTN_SUCCESS
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

            Catch ex As Exception
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
            Dim oMatrix As SAPbouiCOM.Matrix = FormUID.Items.Item("matGtTxtFi").Specific
            Dim oCheckbox As SAPbouiCOM.CheckBox
            Dim oFlag As Boolean = False
            oDT_TxtFileGeneration = New DataTable

            oDT_TxtFileGeneration.Columns.Add("FolderPath", GetType(String))
            oDT_TxtFileGeneration.Columns.Add("DateFrom", GetType(String))
            oDT_TxtFileGeneration.Columns.Add("DateTo", GetType(String))
            oDT_TxtFileGeneration.Columns.Add("OUCodeFrom", GetType(String))
            oDT_TxtFileGeneration.Columns.Add("OUCodeTo", GetType(String))
            oDT_TxtFileGeneration.Columns.Add("Entity", GetType(String))
            oDT_TxtFileGeneration.Columns.Add("EntityDesc", GetType(String))
            Try
                sFuncName = "HeaderValidation()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

                If FormUID.Items.Item("txtFldPath").Specific.string = String.Empty Then
                    p_oSBOApplication.StatusBar.SetText("Folder path should not Empty ....... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    FormUID.ActiveItem = "txtFldPath"
                    Return RTN_ERROR

                ElseIf FormUID.Items.Item("txtDatefrm").Specific.string = String.Empty Then
                    p_oSBOApplication.StatusBar.SetText("Date From is Missing ....... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    FormUID.ActiveItem = "txtDatefrm"
                    Return RTN_ERROR

                ElseIf FormUID.Items.Item("9").Specific.string = String.Empty Then
                    p_oSBOApplication.StatusBar.SetText("Date To is Missing ....... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    FormUID.ActiveItem = "9"
                    Return RTN_ERROR

                ElseIf FormUID.Items.Item("txtOUCdeRn").Specific.string = String.Empty Then
                    p_oSBOApplication.StatusBar.SetText("OU Code From is Missing ....... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    FormUID.ActiveItem = "txtOUCdeRn"
                    Return RTN_ERROR

                ElseIf FormUID.Items.Item("txtOUCdRTO").Specific.string = String.Empty Then
                    p_oSBOApplication.StatusBar.SetText("OU Code To is Missing ....... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    FormUID.ActiveItem = "txtOUCdRTO"
                    Return RTN_ERROR
                End If

                For imjs As Integer = 1 To oMatrix.RowCount
                    oCheckbox = oMatrix.Columns.Item("V_1").Cells.Item(imjs).Specific
                    If oCheckbox.Checked = True Then


                        oDT_TxtFileGeneration.Rows.Add(FormUID.Items.Item("txtFldPath").Specific.string, GetDate(FormUID.Items.Item("txtDatefrm").Specific.string, p_oDICompany), GetDate(FormUID.Items.Item("9").Specific.string, p_oDICompany), _
                                               FormUID.Items.Item("txtOUCdeRn").Specific.string, FormUID.Items.Item("txtOUCdRTO").Specific.string, _
                                               oMatrix.Columns.Item("Col_0").Cells.Item(imjs).Specific.String, oMatrix.Columns.Item("V_0").Cells.Item(imjs).Specific.String)
                    End If

                Next imjs

                If oDT_TxtFileGeneration.Rows.Count = 0 Then
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

            Dim sFuncName As String = String.Empty

            Dim dateValue As DateTime
            Dim DateString As String = String.Empty
            Dim sSQL As String = String.Empty
            Dim oRs As SAPbobsCOM.Recordset
            Dim sDatesep As String

            sFuncName = "GetDate()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            sSQL = "SELECT DateFormat,DateSep FROM OADM"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("SQL " & sSQL, sFuncName)
            oRs.DoQuery(sSQL)


            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Input Date  String  " & sDate, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Date Format  " & oRs.Fields.Item("DateFormat").Value, sFuncName)

            If Not oRs.EoF Then
                sDatesep = oRs.Fields.Item("DateSep").Value

                Select Case oRs.Fields.Item("DateFormat").Value
                    Case 0
                        If Date.TryParseExact(sDate, "dd" & sDatesep & "MM" & sDatesep & "yy", _
                           New CultureInfo("en-US"), _
                           DateTimeStyles.None, _
                           dateValue) Then

                            DateString = dateValue.ToString("yyyyMMdd")
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Case 0 " & DateString, sFuncName)

                        End If
                    Case 1
                        If Date.TryParseExact(sDate, "dd" & sDatesep & "MM" & sDatesep & "yyyy", _
                           New CultureInfo("en-US"), _
                           DateTimeStyles.None, _
                           dateValue) Then

                            DateString = dateValue.ToString("yyyyMMdd")
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Case 1 " & DateString, sFuncName)
                        End If
                    Case 2
                        If Date.TryParseExact(sDate, "MM" & sDatesep & "dd" & sDatesep & "yy", _
                            New CultureInfo("en-US"), _
                            DateTimeStyles.None, _
                            dateValue) Then
                            DateString = dateValue.ToString("yyyyMMdd")
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Case 2 " & DateString, sFuncName)
                        End If
                    Case 3
                        If Date.TryParseExact(sDate, "MM" & sDatesep & "dd" & sDatesep & "yyyy", _
                            New CultureInfo("en-US"), _
                            DateTimeStyles.None, _
                            dateValue) Then
                            DateString = dateValue.ToString("yyyyMMdd")
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Case 3 " & DateString, sFuncName)
                        End If
                    Case 4
                        If Date.TryParseExact(sDate, "yyyy" & sDatesep & "MM" & sDatesep & "dd", _
                            New CultureInfo("en-US"), _
                            DateTimeStyles.None, _
                            dateValue) Then
                            DateString = dateValue.ToString("yyyyMMdd")
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Case 4 " & DateString, sFuncName)
                        End If
                    Case 5
                        If Date.TryParseExact(sDate, "dd" & sDatesep & "MMMM" & sDatesep & "yyyy", _
                            New CultureInfo("en-US"), _
                            DateTimeStyles.None, _
                            dateValue) Then
                            DateString = dateValue.ToString("yyyyMMdd")
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Case 5 " & DateString, sFuncName)
                        End If
                    Case 6
                        If Date.TryParseExact(sDate, "yy" & sDatesep & "MM" & sDatesep & "dd", _
                            New CultureInfo("en-US"), _
                            DateTimeStyles.None, _
                            dateValue) Then
                            DateString = dateValue.ToString("yyyyMMdd")
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Case 6 " & DateString, sFuncName)
                        End If
                    Case Else
                        DateString = dateValue.ToString("yyyyMMdd")
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Else " & DateString, sFuncName)
                End Select

            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS ", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Return Value " & DateString, sFuncName)

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


            'Dim NewCol As DataColumn
            'Dim NewRow As DataRow
            'Dim ColCount As Integer
            'Dim dAmount As Decimal = 0.0

            Dim dAmount As Decimal = 0.0
            sFuncName = "ConvertRecordset()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Fuction ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Data Row Count  " & dtTable.Rows.Count, sFuncName)
            Try
                ''If dtTable.Rows.Count = 0 Then
                ''    For ColCount = 0 To SAPRecordset.Fields.Count - 1
                ''        NewCol = New DataColumn(SAPRecordset.Fields.Item(ColCount).Name)
                ''        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Column Creation  ", sFuncName)
                ''        Try
                ''            dtTable.Columns.Add(NewCol)

                ''        Catch ex As Exception
                ''        End Try
                ''    Next
                ''End If


                Do Until SAPRecordset.EoF

                    ''NewRow = dtTable.NewRow
                    ' ''populate each column in the row we're creating
                    ''For ColCount = 0 To SAPRecordset.Fields.Count - 1

                    ''    NewRow.Item(SAPRecordset.Fields.Item(ColCount).Name) = Convert.ToDecimal(SAPRecordset.Fields.Item(ColCount).Value)

                    ''Next
                    dAmount = FormatNumber(CDbl(SAPRecordset.Fields.Item(5).Value), 3)

                    dtTable.Rows.Add(SAPRecordset.Fields.Item(0).Value, SAPRecordset.Fields.Item(1).Value, SAPRecordset.Fields.Item(2).Value, _
                                     SAPRecordset.Fields.Item(3).Value, SAPRecordset.Fields.Item(4).Value, dAmount)
                    'Add the row to the datatable
                    ''dtTable.Rows.Add(NewRow)


                    SAPRecordset.MoveNext()
                Loop

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS ", sFuncName)

                Return dtTable

            Catch ex As Exception
                sErrDesc = ex.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
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

        Public Sub ShowFileDialog()

            Dim oDialogBox As New FolderBrowserDialog

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
                '' OpenFilePath.FolderBrowserDialog1.Multiselect = False
                ''OpenFilePath.FolderBrowserDialog1.RootFolder = "SAP Business One"
                OpenFilePath.FolderBrowserDialog1.ShowNewFolderButton = True


                If OpenFilePath.FolderBrowserDialog1.ShowDialog = DialogResult.OK Then

                    p_sSelectedFilepath = OpenFilePath.FolderBrowserDialog1.SelectedPath

                    OpenFilePath.OpenFileDialog1.Dispose()
                    OpenFilePath.Close()

                Else
                    p_sSelectedFilepath = String.Empty
                    OpenFilePath.Close()
                    System.Windows.Forms.Application.ExitThread()
                End If


                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Catch ex As Exception
                p_sSelectedFilepath = String.Empty
                sErrDesc = ex.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Finally
            End Try
        End Sub

        Public Function fillopen() As String
            Dim sFuncName As String = String.Empty
            Dim ShowFolderBrowserThread As Threading.Thread
            Try
                sFuncName = "fillopen()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

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

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Catch ex As Exception
                fillopen = String.Empty
                sErrDesc = ex.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            End Try
        End Function

        Public Sub LoadFromXML_AuditFile(ByVal FileName As String, ByVal Sbo_application As SAPbouiCOM.Application)
            Try
                Dim oXmlDoc As New Xml.XmlDocument
                Dim sPath As String
                'sPath = IO.Directory.GetParent(Application.StartupPath).ToString
                sPath = Application.StartupPath.ToString
                'oXmlDoc.Load(sPath & "\AE_FleetMangement\" & FileName)
                oXmlDoc.Load(sPath & "\" & FileName)
                ' MsgBox(Application.StartupPath)

                Sbo_application.LoadBatchActions(oXmlDoc.InnerXml)
            Catch ex As Exception
                MsgBox(ex)
            End Try

        End Sub

#End Region


    End Module
End Namespace



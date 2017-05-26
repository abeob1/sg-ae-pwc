Option Explicit On
Imports System.Xml
Imports System.IO
Imports System.Windows.Forms
Imports System.Globalization
Imports System.Net.Mail
Imports System.Configuration
Imports Microsoft.Office.Interop
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Data



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
                ''  objCompany.LicenseServer = "WIN-D6KRARO05H9:30000"
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

        Public Function MAtrixToDataTable(ByVal oform As SAPbouiCOM.Form, ByRef sErrDesc As String) As DataView

            ' **********************************************************************************
            '   Function    :   GetDataViewFromCSV()
            '   Purpose     :   This function will upload the data from CSV file to Dataview
            '   Parameters  :   ByRef CurrFileToUpload AS String 
            '                       CurrFileToUpload = File Name
            '   Author      :   JOHN
            '   Date        :   MAY 2014 20
            ' **********************************************************************************

            Dim dv As DataView

            Dim sFuncName As String = String.Empty
            sErrDesc = String.Empty
            Dim oMatrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_24").Specific
            Dim oipower As DataTable = Nothing
            Dim sNewOU As String = String.Empty
            Dim sGL As String = String.Empty
            Dim sSQL As String = String.Empty
            Dim orset As SAPbobsCOM.Recordset = Nothing
            Dim oDVNew As DataView = Nothing
            Dim oDVBU As DataView = Nothing
            Dim oDVCom As DataView = Nothing
            Dim oDTError As DataTable = Nothing
            Dim sJVGroup As String = String.Empty

            Try
                sFuncName = "MAtrixToDataTable"
                Console.WriteLine("Starting Function", sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
                orset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'The Datatable to Return
                oipower = New DataTable()
                oDTError = New DataTable()
                oDTError.Columns.Add("OU", GetType(String))
                oDTError.Columns.Add("Error", GetType(String))

                oipower.Columns.Add("GL_Code", GetType(String))
                oipower.Columns.Add("GL_Name", GetType(String))
                oipower.Columns.Add("GL_NameT", GetType(String))
                oipower.Columns.Add("OU_BU_Budget", GetType(String))
                oipower.Columns.Add("Project", GetType(String))
                oipower.Columns.Add("OU", GetType(String))
                oipower.Columns.Add("OUName", GetType(String))
                oipower.Columns.Add("NewOU", GetType(String))
                oipower.Columns.Add("Amount", GetType(Double))
                oipower.Columns.Add("Remarks", GetType(String))
                oipower.Columns.Add("JV", GetType(String))
                oipower.Columns.Add("Base", GetType(String))
                oipower.Columns.Add("EntityCode", GetType(String))
                oipower.Columns.Add("TGL", GetType(String))
                oipower.Columns.Add("UName", GetType(String))
                oipower.Columns.Add("Pass", GetType(String))
                oipower.Columns.Add("JVGroup", GetType(String))
                oipower.Columns.Add("Cat", GetType(String))
                oipower.Columns.Add("BU", GetType(String))
                oipower.Columns.Add("LOS", GetType(String))
                oipower.Columns.Add("OcrCode3", GetType(String))


                ''sSQL = "SELECT T0.[PrcCode], T0.[PrcName], T0.[U_AB_ENTITY], T1.[AcctCode], T1.[AcctName] FROM OPRC T0 join OACT T1 on T0.U_AB_ENTITY = T1.Details WHERE T0.[DimCode] = 3"
                sSQL = "SELECT T0.[PrcCode], T0.[PrcName], T0.[U_AB_ENTITY], T1.[AcctCode], T1.[AcctName] FROM OPRC T0 join OACT T1 on T1.Details  like '%' + T0.U_AB_ENTITY + '%' WHERE T0.[DimCode] = 3"

                orset.DoQuery(sSQL)
                oDVNew = New DataView(ConvertRecordsetToDataTable(orset, sErrDesc))
                If Not String.IsNullOrEmpty(sErrDesc) Then Throw New ArgumentException(sErrDesc)

                sSQL = "SELECT T0.[PrcCode] 'OU' , T0.[U_AB_REPORTCODE] 'BU' , T1.[U_AB_REPORTCODE] 'LOS' FROM OPRC T0 left outer join  OPRC T1 on T0.U_AB_REPORTCODE =  T1.PrcCode"
                orset.DoQuery(sSQL)
                oDVBU = New DataView(ConvertRecordsetToDataTable(orset, sErrDesc))
                If Not String.IsNullOrEmpty(sErrDesc) Then Throw New ArgumentException(sErrDesc)

                sSQL = "SELECT T0.[U_AB_USERCODE], T0.[U_AB_PASSWORD], T0.[U_AB_COMCODE], T0.[U_AB_COMPANYNAME], T0.[U_AB_GROUP], T0.[U_AB_JVCREDIT] , T0.[U_AB_JVDEBIT], T0.[U_AB_NONJVCREDIT], T0.[U_AB_DEPACC], T0.[U_AB_DEPALLACC] FROM [dbo].[@AB_COMPANYDATA]  T0"
                orset.DoQuery(sSQL)
                oDVCom = New DataView(ConvertRecordsetToDataTable(orset, sErrDesc))
                If Not String.IsNullOrEmpty(sErrDesc) Then Throw New ArgumentException(sErrDesc)

                oDVCom.RowFilter = "U_AB_COMCODE='" & p_oDICompany.CompanyDB & "'"

                If oDVCom.Count > 0 Then
                    P_sJV_Debit = oDVCom.Item(0)("U_AB_JVDEBIT").ToString()
                    P_sJV_Credit = oDVCom.Item(0)("U_AB_JVCREDIT").ToString()
                    P_sNonJV_Credit = oDVCom.Item(0)("U_AB_NONJVCREDIT").ToString()
                    P_sDEP_CA = oDVCom.Item(0)("U_AB_DEPACC").ToString()
                    P_sDEP_FAA = oDVCom.Item(0)("U_AB_DEPALLACC").ToString()
                Else
                    sErrDesc = "Pls. define the JV and NON JV GL Accounts in the company setup table "
                    If Not String.IsNullOrEmpty(sErrDesc) Then Throw New ArgumentException(sErrDesc)
                End If

                For imjs As Integer = 1 To oMatrix.RowCount
                    If String.IsNullOrEmpty(oMatrix.Columns.Item("Col_0").Cells.Item(imjs).Specific.string) Then Continue For
                    If Not String.IsNullOrEmpty(oMatrix.Columns.Item("Col_6").Cells.Item(imjs).Specific.string) Then
                        sNewOU = oMatrix.Columns.Item("Col_6").Cells.Item(imjs).Specific.string
                    Else
                        sNewOU = oMatrix.Columns.Item("Col_4").Cells.Item(imjs).Specific.string
                    End If

                    If Not String.IsNullOrEmpty(oMatrix.Columns.Item("V_5").Cells.Item(imjs).Specific.string) Then
                        sGL = oMatrix.Columns.Item("V_5").Cells.Item(imjs).Specific.string
                    Else
                        sGL = oMatrix.Columns.Item("Col_0").Cells.Item(imjs).Specific.string
                    End If

                    oDVNew.RowFilter = "PrcCode='" & oMatrix.Columns.Item("Col_4").Cells.Item(imjs).Specific.string & "'"
                    oDVBU.RowFilter = "OU='" & oMatrix.Columns.Item("Col_4").Cells.Item(imjs).Specific.string & "'"
                    If oDVNew.Count > 0 Then
                        If p_oDICompany.CompanyDB = oDVNew.Item(0)("U_AB_ENTITY").ToString() Then
                            sErrDesc = "Operating Unit " & oMatrix.Columns.Item("Col_4").Cells.Item(imjs).Specific.string & " mapped with the current login entity ...... !"
                            Exit Function
                        End If

                        oDVCom.RowFilter = "U_AB_COMCODE='" & oDVNew.Item(0)("U_AB_ENTITY").ToString() & "'"
                        If oDVCom.Count > 0 Then
                            If oDVCom.Item(0)("U_AB_GROUP").ToString() = "JV" Or oDVCom.Item(0)("U_AB_GROUP").ToString() = "OTHS" Then
                                sJVGroup = "JV"
                            Else
                                sJVGroup = oDVCom.Item(0)("U_AB_GROUP").ToString()
                            End If
                            oipower.Rows.Add(oMatrix.Columns.Item("Col_0").Cells.Item(imjs).Specific.string, oMatrix.Columns.Item("Col_1").Cells.Item(imjs).Specific.string, sGL, oMatrix.Columns.Item("Col_2").Cells.Item(imjs).Specific.string _
                        , oMatrix.Columns.Item("Col_3").Cells.Item(imjs).Specific.string, oMatrix.Columns.Item("Col_4").Cells.Item(imjs).Specific.string, oMatrix.Columns.Item("Col_5").Cells.Item(imjs).Specific.string _
                        , sNewOU, oMatrix.Columns.Item("Col_7").Cells.Item(imjs).Specific.string, oMatrix.Columns.Item("Col_8").Cells.Item(imjs).Specific.string.ToString.Replace("'", "#$%") _
                        , oMatrix.Columns.Item("Col_9").Cells.Item(imjs).Specific.string, oMatrix.Columns.Item("Col_10").Cells.Item(imjs).Specific.string, oDVNew.Item(0)("U_AB_ENTITY").ToString(), oDVNew.Item(0)("AcctCode").ToString() _
                        , oDVCom.Item(0)("U_AB_USERCODE").ToString(), oDVCom.Item(0)("U_AB_PASSWORD").ToString(), sJVGroup, oMatrix.Columns.Item("Col_11").Cells.Item(imjs).Specific.string, oDVBU.Item(0)("BU").ToString(), oDVBU.Item(0)("LOS").ToString() _
                        , oMatrix.Columns.Item("V_4").Cells.Item(imjs).Specific.string)

                            ''          oipower.Rows.Add(oMatrix.Columns.Item("Col_0").Cells.Item(imjs).Specific.string, oMatrix.Columns.Item("Col_1").Cells.Item(imjs).Specific.string, oMatrix.Columns.Item("Col_2").Cells.Item(imjs).Specific.string _
                            '', oMatrix.Columns.Item("Col_3").Cells.Item(imjs).Specific.string, oMatrix.Columns.Item("Col_4").Cells.Item(imjs).Specific.string, oMatrix.Columns.Item("Col_5").Cells.Item(imjs).Specific.string _
                            '', sNewOU, oMatrix.Columns.Item("Col_7").Cells.Item(imjs).Specific.string, oMatrix.Columns.Item("Col_8").Cells.Item(imjs).Specific.string _
                            '', oMatrix.Columns.Item("Col_9").Cells.Item(imjs).Specific.string, oMatrix.Columns.Item("Col_10").Cells.Item(imjs).Specific.string, oDVNew.Item(0)("U_AB_ENTITY").ToString(), oDVNew.Item(0)("AcctCode").ToString() _
                            '', oDVCom.Item(0)("U_AB_USERCODE").ToString(), oDVCom.Item(0)("U_AB_PASSWORD").ToString(), oDVCom.Item(0)("U_AB_GROUP").ToString(), oMatrix.Columns.Item("Col_11").Cells.Item(imjs).Specific.string, oDVBU.Item(0)("BU").ToString(), oDVBU.Item(0)("LOS").ToString())
                        Else
                            oDTError.Rows.Add(oDVNew.Item(0)("U_AB_ENTITY").ToString(), "No Credentials mapped in the Companydata UDT ")
                        End If
                    Else
                        oDTError.Rows.Add(sNewOU, "No Entities are mapped in OACT ")
                    End If
                Next

                If oDTError.Rows.Count > 0 Then
                    Write_TextFileError(oDTError, System.Windows.Forms.Application.StartupPath, sErrDesc)
                    sErrDesc = "Validation Error Occurres ......! "
                    Return Nothing
                End If

                'Console.WriteLine("Del_schema() ", sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS ", sFuncName)
                'Del_schema(p_oCompDef.sInboxDir)
                sErrDesc = String.Empty
                dv = New DataView(oipower)
                If dv.Count = 0 Then
                    sErrDesc = "No Matching records found ......!"
                End If
                Return dv

            Catch ex As Exception
                sErrDesc = ex.Message
                Console.WriteLine("Error occured while reading content of  " & ex.Message, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error occured while reading content of  " & ex.Message, sFuncName)
                Call WriteToLogFile(ex.Message, sFuncName)
                Return Nothing
            Finally

            End Try

        End Function

        Public Function MAtrixToDataTable_E(ByVal oform As SAPbouiCOM.Form, ByRef sErrDesc As String) As DataSet

            ' **********************************************************************************
            '   Function    :   GetDataViewFromCSV()
            '   Purpose     :   This function will upload the data from CSV file to Dataview
            '   Parameters  :   ByRef CurrFileToUpload AS String 
            '                       CurrFileToUpload = File Name
            '   Author      :   JOHN
            '   Date        :   MAY 2014 20
            ' **********************************************************************************

            Dim dv As DataView

            Dim sFuncName As String = String.Empty
            sErrDesc = String.Empty
            Dim oMatrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_24").Specific
            Dim oipower As DataTable = Nothing
            Dim sNewOU As String = String.Empty
            Dim sSQL As String = String.Empty
            Dim orset As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oDVNew As DataView = Nothing
            Dim oDVBU As DataView = Nothing
            Dim oDVCom As DataView = Nothing
            Dim oDTError As DataTable = Nothing
            Dim sGL As String = String.Empty
            Dim oDS As DataSet = Nothing


            Try
                sFuncName = "MAtrixToDataTable"
                Console.WriteLine("Starting Function", sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

                'The Datatable to Return
                oDS = New DataSet()
                oipower = New DataTable()
                oipower = oDS.Tables.Add("JE")
                oDTError = New DataTable()
                oDTError.Columns.Add("OU", GetType(String))
                oDTError.Columns.Add("Error", GetType(String))

                oipower.Columns.Add("GLCode", GetType(String))
                oipower.Columns.Add("GLName", GetType(String)) ' Date
                oipower.Columns.Add("OU_BU_Budget", GetType(String))
                oipower.Columns.Add("Project", GetType(String)) ' Amount
                oipower.Columns.Add("OU", GetType(String))
                oipower.Columns.Add("OUName", GetType(String))
                oipower.Columns.Add("NewOU", GetType(String))
                oipower.Columns.Add("NewGL", GetType(String))
                oipower.Columns.Add("Amount", GetType(Double))
                oipower.Columns.Add("Remarks", GetType(String))
                oipower.Columns.Add("BU", GetType(String))
                oipower.Columns.Add("LOS", GetType(String))
                oipower.Columns.Add("Partner", GetType(String))

                sSQL = "SELECT T0.[PrcCode], T0.[PrcName], T0.[U_AB_ENTITY], T1.[AcctCode], T1.[AcctName] FROM OPRC T0 join OACT T1 on T0.U_AB_ENTITY = T1.Details WHERE T0.[DimCode] = 3"
                orset.DoQuery(sSQL)
                oDVNew = New DataView(ConvertRecordsetToDataTable(orset, sErrDesc))
                If Not String.IsNullOrEmpty(sErrDesc) Then Throw New ArgumentException(sErrDesc)

                sSQL = "SELECT T0.[PrcCode] 'OU' , T0.[U_AB_REPORTCODE] 'BU' , T1.[U_AB_REPORTCODE] 'LOS' FROM OPRC T0 left outer join  OPRC T1 on T0.U_AB_REPORTCODE =  T1.PrcCode"
                orset.DoQuery(sSQL)
                oDVBU = New DataView(ConvertRecordsetToDataTable(orset, sErrDesc))
                If Not String.IsNullOrEmpty(sErrDesc) Then Throw New ArgumentException(sErrDesc)

                For imjs As Integer = 1 To oMatrix.RowCount

                    If String.IsNullOrEmpty(oMatrix.Columns.Item("Col_0").Cells.Item(imjs).Specific.string) Then Continue For
                    If Not String.IsNullOrEmpty(oMatrix.Columns.Item("Col_6").Cells.Item(imjs).Specific.string) Then
                        sNewOU = oMatrix.Columns.Item("Col_6").Cells.Item(imjs).Specific.string
                    Else
                        sNewOU = oMatrix.Columns.Item("Col_4").Cells.Item(imjs).Specific.string
                    End If
                    If Not String.IsNullOrEmpty(oMatrix.Columns.Item("V_5").Cells.Item(imjs).Specific.string) Then
                        sGL = oMatrix.Columns.Item("V_5").Cells.Item(imjs).Specific.string
                    Else
                        sGL = oMatrix.Columns.Item("Col_0").Cells.Item(imjs).Specific.string
                    End If

                    oDVNew.RowFilter = "PrcCode='" & oMatrix.Columns.Item("Col_4").Cells.Item(imjs).Specific.string & "'"
                    oDVBU.RowFilter = "OU='" & oMatrix.Columns.Item("Col_4").Cells.Item(imjs).Specific.string & "'"
                    If oDVNew.Count > 0 Then

                        oipower.Rows.Add(oMatrix.Columns.Item("Col_0").Cells.Item(imjs).Specific.string, oMatrix.Columns.Item("Col_1").Cells.Item(imjs).Specific.string, oMatrix.Columns.Item("Col_2").Cells.Item(imjs).Specific.string _
                    , oMatrix.Columns.Item("Col_3").Cells.Item(imjs).Specific.string, oMatrix.Columns.Item("Col_4").Cells.Item(imjs).Specific.string, oMatrix.Columns.Item("Col_5").Cells.Item(imjs).Specific.string _
                    , sNewOU, sGL, oMatrix.Columns.Item("Col_7").Cells.Item(imjs).Specific.string, oMatrix.Columns.Item("Col_8").Cells.Item(imjs).Specific.string _
                   , oDVBU.Item(0)("BU").ToString(), oDVBU.Item(0)("LOS").ToString(), oMatrix.Columns.Item("V_3").Cells.Item(imjs).Specific.string)

                    Else
                        oDTError.Rows.Add(sNewOU, "No Entities are mapped in OACT ")
                    End If
                Next

                If oDTError.Rows.Count > 0 Then
                    Write_TextFileError(oDTError, System.Windows.Forms.Application.StartupPath, sErrDesc)
                    sErrDesc = "Validation Error Occurres ......! "
                    Return Nothing
                End If

                'Console.WriteLine("Del_schema() ", sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS ", sFuncName)
                'Del_schema(p_oCompDef.sInboxDir)
                sErrDesc = String.Empty

                Return oDS

            Catch ex As Exception
                sErrDesc = ex.Message
                Console.WriteLine("Error occured while reading content of  " & ex.Message, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error occured while reading content of  " & ex.Message, sFuncName)
                Call WriteToLogFile(ex.Message, sFuncName)
                Return Nothing
            Finally

            End Try

        End Function

        Public Function JournalEntry_Posting_JV_Source(ByVal oDVJour As DataView, ByRef oCompany As SAPbobsCOM.Company, ByVal sDate As String, ByRef sRef As String, ByRef sErrDesc As String) As Long

            Dim sFuncName As String = String.Empty
            Dim ival As Integer
            Dim IsError As Boolean
            Dim iErr As Integer = 0
            Dim sErr As String = String.Empty
            Dim sJV As String = String.Empty
            Dim bDebit As Double = 0.0
            Dim bCredit As Double = 0.0

            Dim oRset As SAPbobsCOM.Recordset = Nothing
            Dim sSQL As String = String.Empty
            Dim sPostingDate As String = String.Empty
            Dim oJournalEntry As SAPbobsCOM.JournalEntries = Nothing
            Dim sBU As String = String.Empty
            Dim sOU As String = String.Empty
            Dim sLOS As String = String.Empty
            Dim sOUNAme As String = String.Empty
            Dim sLineRemarks As String = String.Empty
            Dim sNonproject As String = String.Empty
            Dim sProject As String = String.Empty
            Dim sBUC As String = String.Empty
            Dim sOUC As String = String.Empty
            Dim sLOSC As String = String.Empty
            Dim sOUNAmeC As String = String.Empty
            Dim sLineRemarksC As String = String.Empty
            Dim sNonprojectC As String = String.Empty
            Dim sProjectC As String = String.Empty
            Dim sCat As String = String.Empty
            Dim dDate As Date
            Dim bCal As Double = 0
            Dim FCredit As Boolean = False

            Try
                sFuncName = "JournalEntry_Posting"
                Console.WriteLine("Starting Function ", sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)
                oRset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                sSQL = "SELECT T0.[AcctCode], T0.[AcctName] FROM OACT T0"
                oRset.DoQuery(sSQL)
                Dim odt As DataTable = Nothing
                odt = New DataTable()
                odt = ConvertRecordsetToDataTable(oRset, sErrDesc)
                Dim odv As DataView = Nothing
                odv = New DataView(odt)

                oJournalEntry = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
                oJournalEntry.ReferenceDate = DateTime.ParseExact(sDate, "yyyyMMdd", Nothing)
                oJournalEntry.DueDate = DateTime.ParseExact(sDate, "yyyyMMdd", Nothing)
                ''  oJournalEntry.Indicator = "CA"
                dDate = DateTime.ParseExact(sDate, "yyyyMMdd", Nothing)
                oJournalEntry.Memo = Left("Cost Allocation for the month of " & UCase(MonthName(Month(dDate))) & " - " & Year(dDate), 50)
                For Each odr As DataRowView In oDVJour
                    If odr("Amount") > 0 Then

                        bDebit += CDbl(odr("Amount").ToString.Trim)
                        sLOS = odr("LOS").ToString.Trim 'LOS
                        sBU = odr("BU").ToString.Trim 'BU
                        sOU = odr("NewOU").ToString.Trim
                        sProject = odr("Project").ToString.Trim
                        sNonproject = odr("OU_BU_Budget").ToString.Trim 'OU_BU
                        sLineRemarks = odr("Remarks").ToString.Trim.Replace("#$%", "'") 'OU_BU
                        sCat = odr("Cat").ToString.Trim

                        If odr("Cat").ToString().Trim() = "DEP" Then
                            odv.RowFilter = "AcctCode='" & P_sDEP_FAA & "'"
                            If odv.Count = 0 Then
                                Throw New ArgumentException(P_sDEP_FAA & "  - Account Code is missing in Source Entity " & oCompany.CompanyDB & " - " & oCompany.CompanyName)
                            End If
                            oJournalEntry.Lines.AccountCode = P_sDEP_FAA
                        Else
                            odv.RowFilter = "AcctCode='" & Left(odr("GL_Code").ToString.Trim, Len(odr("GL_Code").ToString.Trim) - 4) & "9999" & "'"
                            If odv.Count = 0 Then
                                Throw New ArgumentException(Left(odr("GL_Code").ToString.Trim, Len(odr("GL_Code").ToString.Trim) - 4) & "9999" & "  - Account Code is missing in Source Entity " & oCompany.CompanyDB & " - " & oCompany.CompanyName)
                            End If
                            oJournalEntry.Lines.AccountCode = Left(odr("GL_Code").ToString.Trim, Len(odr("GL_Code").ToString.Trim) - 4) & "9999"
                        End If

                        oJournalEntry.Lines.Credit = CDbl(odr("Amount").ToString.Trim)
                        '' oJournalEntry.Lines.CostingCode = odr(18).ToString.Trim
                        '' oJournalEntry.Lines.CostingCode3 = odr(10).ToString.Trim
                        Select Case odr("Cat").ToString.Trim()
                            Case "AP"
                                oJournalEntry.Lines.UserFields.Fields.Item("U_AB_AP").Value = odr("Base").ToString.Trim
                            Case "CN"
                                oJournalEntry.Lines.UserFields.Fields.Item("U_AB_APCN").Value = odr("Base").ToString.Trim
                        End Select
                        oJournalEntry.Lines.UserFields.Fields.Item("U_AB_JV").Value = odr("JV").ToString.Trim
                        If Not String.IsNullOrEmpty(odr("LOS").ToString.Trim) Then
                            oJournalEntry.Lines.CostingCode = odr("LOS").ToString.Trim 'LOS
                        End If
                        If Not String.IsNullOrEmpty(odr("BU").ToString.Trim) Then
                            oJournalEntry.Lines.CostingCode2 = odr("BU").ToString.Trim 'BU
                        End If

                        If Not String.IsNullOrEmpty(odr("NewOU").ToString.Trim) Then
                            oJournalEntry.Lines.CostingCode3 = odr("NewOU").ToString.Trim 'OU
                        End If
                        If Not String.IsNullOrEmpty(odr("Project").ToString.Trim) Then
                            oJournalEntry.Lines.ProjectCode = odr("Project").ToString.Trim 'Project
                        End If
                        oJournalEntry.Lines.UserFields.Fields.Item("U_AB_NONPROJECT").Value = odr("OU_BU_Budget").ToString.Trim 'OU_BU
                        oJournalEntry.Lines.UserFields.Fields.Item("U_AB_REMARKS").Value = odr("Remarks").ToString.Trim.Replace("#$%", "'") 'OU_BU
                        oJournalEntry.Lines.UserFields.Fields.Item("U_AB_OcrCode3").Value = odr("OcrCode3").ToString.Trim '

                    Else
                        If odr("Cat").ToString().Trim() = "DEP" Then
                            odv.RowFilter = "AcctCode='" & P_sDEP_FAA & "'"
                            If odv.Count = 0 Then
                                Throw New ArgumentException(P_sDEP_FAA & "  - Account Code is missing in Source Entity " & oCompany.CompanyDB & " - " & oCompany.CompanyName)
                            End If

                            oJournalEntry.Lines.AccountCode = P_sDEP_FAA
                        Else
                            odv.RowFilter = "AcctCode='" & Left(odr("GL_Code").ToString.Trim, Len(odr("GL_Code").ToString.Trim) - 4) & "9999" & "'"
                            If odv.Count = 0 Then
                                Throw New ArgumentException(Left(odr("GL_Code").ToString.Trim, Len(odr("GL_Code").ToString.Trim) - 4) & "9999" & "  - Account Code is missing in Source Entity " & oCompany.CompanyDB & " - " & oCompany.CompanyName)
                            End If

                            oJournalEntry.Lines.AccountCode = Left(odr("GL_Code").ToString.Trim, Len(odr("GL_Code").ToString.Trim) - 4) & "9999"
                        End If
                        '' oJournalEntry.Lines.AccountCode = Left(odr("GL_Code").ToString.Trim, Len(odr("GL_Code").ToString.Trim) - 4) & "9999"
                        oJournalEntry.Lines.Debit = Math.Abs(CDbl(odr("Amount").ToString.Trim))
                        Select Case odr("Cat").ToString.Trim()
                            Case "AP"
                                oJournalEntry.Lines.UserFields.Fields.Item("U_AB_AP").Value = odr("Base").ToString.Trim
                            Case "CN"
                                oJournalEntry.Lines.UserFields.Fields.Item("U_AB_APCN").Value = odr("Base").ToString.Trim
                        End Select
                        oJournalEntry.Lines.UserFields.Fields.Item("U_AB_JV").Value = odr("JV").ToString.Trim
                        If Not String.IsNullOrEmpty(odr("LOS").ToString.Trim) Then
                            oJournalEntry.Lines.CostingCode = odr("LOS").ToString.Trim 'LOS
                        End If

                        If Not String.IsNullOrEmpty(odr("BU").ToString.Trim) Then
                            oJournalEntry.Lines.CostingCode2 = odr("BU").ToString.Trim 'BU
                        End If

                        If Not String.IsNullOrEmpty(odr("NewOU").ToString.Trim) Then
                            oJournalEntry.Lines.CostingCode3 = odr("NewOU").ToString.Trim 'OU
                        End If
                        If Not String.IsNullOrEmpty(odr("Project").ToString.Trim) Then
                            oJournalEntry.Lines.ProjectCode = odr("Project").ToString.Trim 'Project
                        End If
                        oJournalEntry.Lines.UserFields.Fields.Item("U_AB_NONPROJECT").Value = odr("OU_BU_Budget").ToString.Trim 'OU_BU
                        oJournalEntry.Lines.UserFields.Fields.Item("U_AB_REMARKS").Value = odr("Remarks").ToString.Trim.Replace("#$%", "'") 'OU_BU
                        oJournalEntry.Lines.UserFields.Fields.Item("U_AB_OcrCode3").Value = odr("OcrCode3").ToString.Trim '
                        '' oJournalEntry.Lines.Add()

                        bCredit += Math.Abs(CDbl(odr("Amount").ToString.Trim))
                        sLOSC = odr("LOS").ToString.Trim 'LOS
                        sBUC = odr("BU").ToString.Trim 'BU
                        sOUC = odr("NewOU").ToString.Trim
                        sProjectC = odr("Project").ToString.Trim
                        sNonprojectC = odr("OU_BU_Budget").ToString.Trim 'OU_BU
                        sLineRemarksC = odr("Remarks").ToString.Trim.Replace("#$%", "'") 'OU_BU
                    End If
                    oJournalEntry.Lines.Add()
                Next

                If bDebit > 0 Then
                    If sCat = "DEP" Then
                        odv.RowFilter = "AcctCode='" & P_sDEP_CA & "'"
                        If odv.Count = 0 Then
                            Throw New ArgumentException(P_sDEP_CA & "  - Account Code is missing in Source Entity " & oCompany.CompanyDB & " - " & oCompany.CompanyName)
                        End If
                        oJournalEntry.Lines.AccountCode = P_sDEP_CA  ''"13161300"
                    Else
                        odv.RowFilter = "AcctCode='" & P_sJV_Debit & "'"
                        If odv.Count = 0 Then
                            Throw New ArgumentException(P_sJV_Debit & "  - Account Code is missing in Source Entity " & oCompany.CompanyDB & " - " & oCompany.CompanyName)
                        End If
                        oJournalEntry.Lines.AccountCode = P_sJV_Debit ''"13161400"
                    End If
                    If bCredit > 0 Then
                        bCal = bDebit - bCredit
                        If bCal > 0 Then
                            oJournalEntry.Lines.Debit = bCal
                        Else
                            oJournalEntry.Lines.Credit = Math.Abs(bCal)
                        End If
                        FCredit = True
                    Else
                        oJournalEntry.Lines.Debit = bDebit
                    End If

                    If Not String.IsNullOrEmpty(sLOS) Then
                        oJournalEntry.Lines.CostingCode = sLOS 'LOS
                    End If

                    If Not String.IsNullOrEmpty(sBU) Then
                        oJournalEntry.Lines.CostingCode2 = sBU 'BU
                    End If

                    If Not String.IsNullOrEmpty(sOU) Then
                        oJournalEntry.Lines.CostingCode3 = sOU  'OU
                    End If
                    If Not String.IsNullOrEmpty(sProject) Then
                        oJournalEntry.Lines.ProjectCode = sProject  'Project
                    End If
                    oJournalEntry.Lines.UserFields.Fields.Item("U_AB_NONPROJECT").Value = sNonproject  'OU_BU
                    oJournalEntry.Lines.UserFields.Fields.Item("U_AB_REMARKS").Value = sLineRemarks  'OU_BU
                    oJournalEntry.Lines.Add()
                End If

                If bCredit > 0 And FCredit = False Then
                    If sCat = "DEP" Then
                        odv.RowFilter = "AcctCode='" & P_sDEP_CA & "'"
                        If odv.Count = 0 Then
                            Throw New ArgumentException(P_sDEP_CA & "  - Account Code is missing in Source Entity " & oCompany.CompanyDB & " - " & oCompany.CompanyName)
                        End If
                        oJournalEntry.Lines.AccountCode = P_sDEP_CA  ''"13161300"
                    Else
                        odv.RowFilter = "AcctCode='" & P_sJV_Debit & "'"
                        If odv.Count = 0 Then
                            Throw New ArgumentException(P_sJV_Debit & "  - Account Code is missing in Source Entity " & oCompany.CompanyDB & " - " & oCompany.CompanyName)
                        End If
                        oJournalEntry.Lines.AccountCode = P_sJV_Debit ''"13161400"
                    End If
                    '' oJournalEntry.Lines.AccountCode = P_sJV_Debit ''"13161400"
                    oJournalEntry.Lines.Credit = bCredit
                    If Not String.IsNullOrEmpty(sLOSC) Then
                        oJournalEntry.Lines.CostingCode = sLOSC  'LOS
                    End If

                    If Not String.IsNullOrEmpty(sBUC) Then
                        oJournalEntry.Lines.CostingCode2 = sBUC 'BU
                    End If

                    If Not String.IsNullOrEmpty(sOUC) Then
                        oJournalEntry.Lines.CostingCode3 = sOUC 'OU
                    End If
                    If Not String.IsNullOrEmpty(sProjectC) Then
                        oJournalEntry.Lines.ProjectCode = sProjectC  'Project
                    End If
                    oJournalEntry.Lines.UserFields.Fields.Item("U_AB_NONPROJECT").Value = sNonprojectC  'OU_BU
                    oJournalEntry.Lines.UserFields.Fields.Item("U_AB_REMARKS").Value = sLineRemarksC  'OU_BU
                    oJournalEntry.Lines.Add()
                Else
                    FCredit = False
                End If

                Console.WriteLine("Attempting to Add the Journal Entry ", sFuncName)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add the Journal Entry", sFuncName)
                oCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
                oJournalEntry.SaveXML(System.Windows.Forms.Application.StartupPath & "\JVSource.xml")
                ival = oJournalEntry.Add()

                If ival <> 0 Then
                    IsError = True
                    oCompany.GetLastError(iErr, sErr)
                    Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
                    Console.WriteLine("Completed with ERROR ", sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
                    JournalEntry_Posting_JV_Source = RTN_ERROR
                    Throw New ArgumentException(sErr)
                End If

                Console.WriteLine("Completed with SUCCESS", sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS ", sFuncName)
                oCompany.GetNewObjectCode(sRef)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Journal Entry DocEntry  " & sRef, sFuncName)

                JournalEntry_Posting_JV_Source = RTN_SUCCESS
                sErrDesc = String.Empty
            Catch ex As Exception
                sErrDesc = ex.Message
                Call WriteToLogFile(ex.Message, sFuncName)
                Console.WriteLine("Completed with ERROR ", sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & ex.Message, sFuncName)
                JournalEntry_Posting_JV_Source = RTN_ERROR
                Exit Function
            End Try

        End Function

        Public Function JournalEntry_Posting_JV_Target(ByVal oDVJour As DataView, ByRef oCompany As SAPbobsCOM.Company, ByVal sDate As String, ByVal sRef As String, ByRef sErrDesc As String) As Long

            Dim sFuncName As String = String.Empty
            Dim ival As Integer
            Dim IsError As Boolean
            Dim iErr As Integer = 0
            Dim sErr As String = String.Empty
            Dim sJV As String = String.Empty

            Dim bDebit As Double = 0.0
            Dim bCredit As Double = 0.0

            Dim sBU As String = String.Empty
            Dim sOU As String = String.Empty
            Dim sLOS As String = String.Empty
            Dim sOUNAme As String = String.Empty
            Dim sLineRemarks As String = String.Empty
            Dim sNonproject As String = String.Empty
            Dim sProject As String = String.Empty
            Dim sBUC As String = String.Empty
            Dim sOUC As String = String.Empty
            Dim sLOSC As String = String.Empty
            Dim sOUNAmeC As String = String.Empty
            Dim sLineRemarksC As String = String.Empty
            Dim sNonprojectC As String = String.Empty
            Dim sProjectC As String = String.Empty
            Dim dDate As Date
            Dim oRset As SAPbobsCOM.Recordset = Nothing
            Dim sSQL As String = String.Empty
            Dim sPostingDate As String = String.Empty
            Dim oJournalEntry As SAPbobsCOM.JournalVouchers = Nothing
            Dim oDTJE As DataTable = Nothing
            Dim sAccCode As String = Nothing
            Dim FCredit As Boolean = False
            Dim dcal As Double = 0

            Try
                sFuncName = "JournalEntry_Posting_JV_Target"
                Console.WriteLine("Starting Function ", sFuncName)
                oDTJE = New DataTable()
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)
                oRset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oJournalEntry = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalVouchers)
                oDTJE.Columns.Add("AccountCode", GetType(String))
                oDTJE.Columns.Add("Debit", GetType(Double))
                oDTJE.Columns.Add("Credit", GetType(Double))
                oDTJE.Columns.Add("CostingCode", GetType(String))
                oDTJE.Columns.Add("CostingCode2", GetType(String))
                oDTJE.Columns.Add("CostingCode3", GetType(String))
                oDTJE.Columns.Add("CostingCode4", GetType(String))
                oDTJE.Columns.Add("LineMemo", GetType(String))
                dDate = DateTime.ParseExact(sDate, "yyyyMMdd", Nothing)

                For Each odr As DataRowView In oDVJour
                    If odr("Amount") > 0 Then
                        If odr("Cat").ToString().Trim() = "DEP" Then
                            sAccCode = P_sDEP_FAA
                        Else
                            sAccCode = odr("GL_NameT").ToString.Trim
                        End If

                        oDTJE.Rows.Add(sAccCode, CDbl(odr("Amount").ToString.Trim), 0, odr("LOS").ToString.Trim, odr("BU").ToString.Trim, odr("NewOU").ToString.Trim, odr("Project").ToString.Trim, odr("Remarks").ToString.Trim.Replace("#$%", "'"))
                        bCredit += CDbl(odr("Amount").ToString.Trim)
                        sLOSC = odr("LOS").ToString.Trim
                        sBUC = odr("BU").ToString.Trim
                        sOUC = odr("NewOU").ToString.Trim
                        sProjectC = odr("Project").ToString.Trim
                        sLineRemarksC = odr("Remarks").ToString.Trim.Replace("#$%", "'")
                    Else
                        If odr("Cat").ToString().Trim() = "DEP" Then
                            sAccCode = P_sDEP_FAA
                        Else
                            sAccCode = odr("GL_NameT").ToString.Trim
                        End If
                        oDTJE.Rows.Add(sAccCode, 0, Math.Abs(CDbl(odr("Amount").ToString.Trim)), odr("LOS").ToString.Trim, odr("BU").ToString.Trim, odr("NewOU").ToString.Trim, odr("Project").ToString.Trim, odr("Remarks").ToString.Trim.Replace("#$%", "'"))
                        'oJournalEntry.JournalEntries.Lines.AccountCode = odr("GL_NameT").ToString.Trim
                        bDebit += Math.Abs(CDbl(odr("Amount").ToString.Trim))
                        sLOS = odr("LOS").ToString.Trim 'LOS
                        sBU = odr("BU").ToString.Trim 'BU
                        sOU = odr("NewOU").ToString.Trim 'OU
                        sProject = odr("Project").ToString.Trim 'Project
                        sLineRemarks = odr("Remarks").ToString.Trim.Replace("#$%", "'") 'Project
                    End If
                Next

                If bDebit > 0 Then
                    If bCredit > 0 Then
                        dcal = bDebit - bCredit
                        If dcal > 0 Then
                            oDTJE.Rows.Add(P_sJV_Credit, dcal, 0, sLOS, sBU, sOU, sProject, Left("Cost Allocation for the month of " & UCase(MonthName(Month(dDate))) & " - " & Year(dDate), 50))
                        Else
                            oDTJE.Rows.Add(P_sJV_Credit, 0, Math.Abs(dcal), sLOS, sBU, sOU, sProject, Left("Cost Allocation for the month of " & UCase(MonthName(Month(dDate))) & " - " & Year(dDate), 50))
                        End If

                        FCredit = True
                    Else
                        oDTJE.Rows.Add(P_sJV_Credit, bDebit, 0, sLOS, sBU, sOU, sProject, sLineRemarks)
                    End If
                End If

                If bCredit > 0 And FCredit = False Then
                    oDTJE.Rows.Add(P_sJV_Credit, 0, bCredit, sLOSC, sBUC, sOUC, sProjectC, sLineRemarksC)
                Else
                    FCredit = False
                End If

                sSQL = "SELECT T0.[AcctCode], T0.[AcctName] FROM OACT T0"
                oRset.DoQuery(sSQL)
                Dim odt As DataTable = Nothing
                odt = New DataTable()
                odt = ConvertRecordsetToDataTable(oRset, sErrDesc)
                Dim odv As DataView = Nothing
                odv = New DataView(odt)

                oJournalEntry.JournalEntries.ReferenceDate = DateTime.ParseExact(sDate, "yyyyMMdd", Nothing)
                oJournalEntry.JournalEntries.DueDate = DateTime.ParseExact(sDate, "yyyyMMdd", Nothing)
                oJournalEntry.JournalEntries.Memo = Left("Cost Allocation for the month of " & UCase(MonthName(Month(dDate))) & " - " & Year(dDate), 50)
                oJournalEntry.JournalEntries.Reference3 = sRef
                For Each odr As DataRow In oDTJE.Rows
                    odv.RowFilter = "AcctCode='" & odr("AccountCode").ToString.Trim & "'"
                    If odv.Count = 0 Then
                        Throw New ArgumentException(odr("AccountCode").ToString.Trim & "  - Account Code is missing in Traget Entity " & oCompany.CompanyDB & " - " & oCompany.CompanyName)
                    End If

                    oJournalEntry.JournalEntries.Lines.AccountCode = odr("AccountCode").ToString.Trim
                    oJournalEntry.JournalEntries.Lines.Debit = CDbl(odr("Debit").ToString.Trim)
                    oJournalEntry.JournalEntries.Lines.Credit = CDbl(odr("Credit").ToString.Trim)
                    If Not String.IsNullOrEmpty(odr("CostingCode").ToString.Trim) Then
                        oJournalEntry.JournalEntries.Lines.CostingCode = odr("CostingCode").ToString.Trim 'LOS
                    End If
                    If Not String.IsNullOrEmpty(odr("CostingCode2").ToString.Trim) Then
                        oJournalEntry.JournalEntries.Lines.CostingCode2 = odr("CostingCode2").ToString.Trim 'BU
                    End If
                    If Not String.IsNullOrEmpty(odr("CostingCode3").ToString.Trim) Then
                        oJournalEntry.JournalEntries.Lines.CostingCode3 = odr("CostingCode3").ToString.Trim 'OU
                    End If
                    If Not String.IsNullOrEmpty(odr("CostingCode4").ToString.Trim) Then
                        oJournalEntry.JournalEntries.Lines.ProjectCode = odr("CostingCode4").ToString.Trim 'Project
                    End If
                    oJournalEntry.JournalEntries.Lines.LineMemo = odr("LineMemo").ToString.Trim
                    oJournalEntry.JournalEntries.Lines.Add()
                Next

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add the Journal Entry", sFuncName)
                ival = oJournalEntry.Add()

                If ival <> 0 Then
                    IsError = True
                    oCompany.GetLastError(iErr, sErr)
                    Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
                    Console.WriteLine("Completed with ERROR ", sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
                    JournalEntry_Posting_JV_Target = RTN_ERROR
                    Throw New ArgumentException(sErr)
                End If

                Console.WriteLine("Completed with SUCCESS", sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS ", sFuncName)
                oCompany.GetNewObjectCode(sJV)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Journal Entry DocEntry  " & sJV, sFuncName)

                JournalEntry_Posting_JV_Target = RTN_SUCCESS
                sErrDesc = String.Empty
            Catch ex As Exception
                sErrDesc = ex.Message
                Call WriteToLogFile(ex.Message, sFuncName)
                Console.WriteLine("Completed with ERROR ", sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & ex.Message, sFuncName)
                JournalEntry_Posting_JV_Target = RTN_ERROR
                Exit Function
            End Try

        End Function

        ''Public Function JournalEntry_Posting_NONJV_Source(ByVal oDVJour As DataView, ByRef oCompany As SAPbobsCOM.Company, ByVal sDate As String, ByRef sRef As String, ByRef sErrDesc As String) As Long

        ''    Dim sFuncName As String = String.Empty
        ''    Dim ival As Integer
        ''    Dim IsError As Boolean
        ''    Dim iErr As Integer = 0
        ''    Dim sErr As String = String.Empty
        ''    Dim sJV As String = String.Empty

        ''    Dim oRset As SAPbobsCOM.Recordset = Nothing
        ''    Dim sSQL As String = String.Empty
        ''    Dim sPostingDate As String = String.Empty
        ''    Dim oJournalEntry As SAPbobsCOM.JournalVouchers = Nothing
        ''    Dim bCredit, bDebit As Double
        ''    Dim sBU As String = String.Empty
        ''    Dim sOU As String = String.Empty
        ''    Dim sLOS As String = String.Empty
        ''    Dim sOUNAme As String = String.Empty
        ''    Dim sLineRemarks As String = String.Empty
        ''    Dim sNonproject As String = String.Empty
        ''    Dim sProject As String = String.Empty
        ''    Dim sBUC As String = String.Empty
        ''    Dim sOUC As String = String.Empty
        ''    Dim sLOSC As String = String.Empty
        ''    Dim sOUNAmeC As String = String.Empty
        ''    Dim sLineRemarksC As String = String.Empty
        ''    Dim sNonprojectC As String = String.Empty
        ''    Dim sProjectC As String = String.Empty
        ''    Dim sGL As String = String.Empty
        ''    Dim sCat As String = String.Empty
        ''    Dim sBase As String = String.Empty
        ''    Dim sCatC As String = String.Empty
        ''    Dim sBaseC As String = String.Empty
        ''    Dim sJC As String = String.Empty
        ''    Dim dDate As Date
        ''    Dim dCal As Double = 0
        ''    Dim FCredit As Boolean = False
        ''    Try
        ''        sFuncName = "JournalEntry_Posting_NONJV_Source"
        ''        Console.WriteLine("Starting Function ", sFuncName)
        ''        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)
        ''        oRset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        ''        oJournalEntry = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalVouchers)
        ''        oJournalEntry.JournalEntries.ReferenceDate = DateTime.ParseExact(sDate, "yyyyMMdd", Nothing)
        ''        oJournalEntry.JournalEntries.DueDate = DateTime.ParseExact(sDate, "yyyyMMdd", Nothing)
        ''        dDate = DateTime.ParseExact(sDate, "yyyyMMdd", Nothing)
        ''        oJournalEntry.JournalEntries.Memo = Left("Cost Allocation for the month of " & UCase(MonthName(Month(dDate))) & " - " & Year(dDate), 50)
        ''        ''  oJournalEntry.Indicator = "CA"

        ''        For Each odr As DataRowView In oDVJour
        ''            ''  If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug1(odr("GL_Code") & "  " & odr("GL_NameT") & "  " & odr("OU") & "  " & odr("Amount") & "  " & odr("TGL"), "JournalEntry_Posting_NONJV_Source Inside")
        ''            sGL = odr("TGL").ToString.Trim()
        ''            If odr("Amount") > 0 Then
        ''                ''  oJournalEntry.Lines.AccountCode = odr("TGL").ToString.Trim()
        ''                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug1("Debit", "")

        ''                bDebit += CDbl(odr("Amount").ToString.Trim)
        ''                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug1("Debit Amount " & bDebit, "")
        ''                sCat = odr("Cat").ToString.Trim()
        ''                sBase = odr("Base").ToString.Trim()
        ''                sJV = odr("JV").ToString.Trim
        ''                sLOS = odr("LOS").ToString.Trim
        ''                sBU = odr("BU").ToString.Trim
        ''                sOU = odr("NewOU").ToString.Trim
        ''                sProject = odr("Project").ToString.Trim
        ''                sNonproject = odr("OU_BU_Budget").ToString.Trim
        ''                sLineRemarks = odr("Remarks").ToString.Trim.Replace("#$%", "'")

        ''                If sCat = "DEP" Then
        ''                    oJournalEntry.JournalEntries.Lines.AccountCode = P_sDEP_FAA
        ''                Else
        ''                    oJournalEntry.JournalEntries.Lines.AccountCode = Left(odr("GL_NameT").ToString.Trim, Len(odr("GL_NameT").ToString.Trim) - 4) & "9999"
        ''                End If
        ''                oJournalEntry.JournalEntries.Lines.Credit = CDbl(odr("Amount").ToString.Trim)
        ''                Select Case odr("Cat").ToString.Trim()
        ''                    Case "AP"
        ''                        oJournalEntry.JournalEntries.Lines.UserFields.Fields.Item("U_AB_AP").Value = odr("Base").ToString.Trim
        ''                    Case "CN"
        ''                        oJournalEntry.JournalEntries.Lines.UserFields.Fields.Item("U_AB_APCN").Value = odr("Base").ToString.Trim
        ''                End Select
        ''                oJournalEntry.JournalEntries.Lines.UserFields.Fields.Item("U_AB_JV").Value = odr("JV").ToString.Trim
        ''                If Not String.IsNullOrEmpty(odr("LOS").ToString.Trim) Then
        ''                    oJournalEntry.JournalEntries.Lines.CostingCode = odr("LOS").ToString.Trim 'LOS
        ''                End If

        ''                If Not String.IsNullOrEmpty(odr("BU").ToString.Trim) Then
        ''                    oJournalEntry.JournalEntries.Lines.CostingCode2 = odr("BU").ToString.Trim 'BU
        ''                End If

        ''                If Not String.IsNullOrEmpty(odr("NewOU").ToString.Trim) Then
        ''                    oJournalEntry.JournalEntries.Lines.CostingCode3 = odr("NewOU").ToString.Trim 'OU
        ''                End If
        ''                If Not String.IsNullOrEmpty(odr("Project").ToString.Trim) Then
        ''                    oJournalEntry.JournalEntries.Lines.ProjectCode = odr("Project").ToString.Trim 'Project
        ''                End If
        ''                oJournalEntry.JournalEntries.Lines.UserFields.Fields.Item("U_AB_NONPROJECT").Value = odr("OU_BU_Budget").ToString.Trim 'OU_BU
        ''                oJournalEntry.JournalEntries.Lines.UserFields.Fields.Item("U_AB_REMARKS").Value = odr("Remarks").ToString.Trim.Replace("#$%", "'") 'OU_BU
        ''                oJournalEntry.JournalEntries.Lines.UserFields.Fields.Item("U_AB_OcrCode3").Value = odr("OcrCode3").ToString.Trim '
        ''                ''oJournalEntry.Lines.UserFields.Fields.Item("U_AB_NONPROJECT").Value = odr("OU_BU_Budget").ToString.Trim 'OU_BU
        ''                '' oJournalEntry.Lines.CostingCode = odr(18).ToString.Trim
        ''                '' oJournalEntry.Lines.CostingCode3 = odr(10).ToString.Trim
        ''            Else
        ''                ''  If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug1("Credit", "")

        ''                bCredit += Math.Abs(CDbl(odr("Amount").ToString.Trim))
        ''                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug1("Credit Amount " & bCredit, "")
        ''                sCatC = odr("Cat").ToString.Trim()
        ''                sBaseC = odr("Base").ToString.Trim
        ''                sJC = odr("JV").ToString.Trim
        ''                sLOSC = odr("LOS").ToString.Trim
        ''                sBUC = odr("BU").ToString.Trim
        ''                sOUC = odr("NewOU").ToString.Trim
        ''                sProjectC = odr("Project").ToString.Trim
        ''                sNonprojectC = odr("OU_BU_Budget").ToString.Trim
        ''                sLineRemarksC = odr("Remarks").ToString.Trim.Replace("#$%", "'")

        ''                If sCat = "DEP" Then
        ''                    oJournalEntry.JournalEntries.Lines.AccountCode = P_sDEP_FAA
        ''                Else
        ''                    oJournalEntry.JournalEntries.Lines.AccountCode = Left(odr("GL_NameT").ToString.Trim, Len(odr("GL_NameT").ToString.Trim) - 4) & "9999"
        ''                End If
        ''                '' oJournalEntry.Lines.AccountCode = Left(odr("GL_NameT").ToString.Trim, Len(odr("GL_NameT").ToString.Trim) - 4) & "9999"
        ''                oJournalEntry.JournalEntries.Lines.Debit = Math.Abs(CDbl(odr("Amount").ToString.Trim))
        ''                Select Case odr("Cat").ToString.Trim()
        ''                    Case "AP"
        ''                        oJournalEntry.JournalEntries.Lines.UserFields.Fields.Item("U_AB_AP").Value = odr("Base").ToString.Trim
        ''                    Case "CN"
        ''                        oJournalEntry.JournalEntries.Lines.UserFields.Fields.Item("U_AB_APCN").Value = odr("Base").ToString.Trim
        ''                End Select
        ''                oJournalEntry.JournalEntries.Lines.UserFields.Fields.Item("U_AB_JV").Value = odr("JV").ToString.Trim
        ''                If Not String.IsNullOrEmpty(odr("LOS").ToString.Trim) Then
        ''                    oJournalEntry.JournalEntries.Lines.CostingCode = odr("LOS").ToString.Trim 'LOS
        ''                End If

        ''                If Not String.IsNullOrEmpty(odr("BU").ToString.Trim) Then
        ''                    oJournalEntry.JournalEntries.Lines.CostingCode2 = odr("BU").ToString.Trim 'BU
        ''                End If

        ''                If Not String.IsNullOrEmpty(odr("NewOU").ToString.Trim) Then
        ''                    oJournalEntry.JournalEntries.Lines.CostingCode3 = odr("NewOU").ToString.Trim 'OU
        ''                End If
        ''                If Not String.IsNullOrEmpty(odr("Project").ToString.Trim) Then
        ''                    oJournalEntry.JournalEntries.Lines.ProjectCode = odr("Project").ToString.Trim 'Project
        ''                End If
        ''                oJournalEntry.JournalEntries.Lines.UserFields.Fields.Item("U_AB_NONPROJECT").Value = odr("OU_BU_Budget").ToString.Trim 'OU_BU
        ''                oJournalEntry.JournalEntries.Lines.UserFields.Fields.Item("U_AB_REMARKS").Value = odr("Remarks").ToString.Trim.Replace("#$%", "'") 'OU_BU
        ''                oJournalEntry.JournalEntries.Lines.UserFields.Fields.Item("U_AB_OcrCode3").Value = odr("OcrCode3").ToString.Trim '
        ''                ''  oJournalEntry.Lines.UserFields.Fields.Item("U_AB_NONPROJECT").Value = odr("OU_BU_Budget").ToString.Trim 'OU_BU
        ''                '' oJournalEntry.Lines.CostingCode = odr(18).ToString.Trim
        ''                '' oJournalEntry.Lines.CostingCode3 = odr(10).ToString.Trim
        ''            End If
        ''            oJournalEntry.JournalEntries.Lines.Add()
        ''        Next

        ''        If bDebit > 0 Then
        ''            oJournalEntry.JournalEntries.Lines.AccountCode = sGL
        ''            If bCredit > 0 Then
        ''                dCal = bDebit - bCredit
        ''                If dCal > 0 Then
        ''                    oJournalEntry.JournalEntries.Lines.Debit = dCal
        ''                Else
        ''                    oJournalEntry.JournalEntries.Lines.Credit = Math.Abs(dCal)
        ''                End If
        ''                FCredit = True
        ''            Else
        ''                oJournalEntry.JournalEntries.Lines.Debit = bDebit
        ''            End If
        ''            Select Case sCat
        ''                Case "AP"
        ''                    oJournalEntry.JournalEntries.Lines.UserFields.Fields.Item("U_AB_AP").Value = sBase
        ''                Case "CN"
        ''                    oJournalEntry.JournalEntries.Lines.UserFields.Fields.Item("U_AB_APCN").Value = sBase
        ''            End Select
        ''            oJournalEntry.JournalEntries.Lines.UserFields.Fields.Item("U_AB_JV").Value = sJV
        ''            If Not String.IsNullOrEmpty(sLOS) Then
        ''                oJournalEntry.JournalEntries.Lines.CostingCode = sLOS 'LOS
        ''            End If

        ''            If Not String.IsNullOrEmpty(sBU) Then
        ''                oJournalEntry.JournalEntries.Lines.CostingCode2 = sBU 'BU
        ''            End If

        ''            If Not String.IsNullOrEmpty(sOU) Then
        ''                oJournalEntry.JournalEntries.Lines.CostingCode3 = sOU 'OU
        ''            End If
        ''            If Not String.IsNullOrEmpty(sProject) Then
        ''                oJournalEntry.JournalEntries.Lines.ProjectCode = sProject  'Project
        ''            End If
        ''            oJournalEntry.JournalEntries.Lines.UserFields.Fields.Item("U_AB_NONPROJECT").Value = sNonproject  'OU_BU
        ''            oJournalEntry.JournalEntries.Lines.UserFields.Fields.Item("U_AB_REMARKS").Value = sLineRemarks  'OU_BU
        ''            ''oJournalEntry.Lines.UserFields.Fields.Item("U_AB_NONPROJECT").Value = odr("OU_BU_Budget").ToString.Trim 'OU_BU
        ''            oJournalEntry.JournalEntries.Lines.Add()
        ''        End If

        ''        If bCredit > 0 And FCredit = False Then

        ''            oJournalEntry.JournalEntries.Lines.AccountCode = sGL
        ''            oJournalEntry.JournalEntries.Lines.Credit = bCredit
        ''            Select Case sCatC
        ''                Case "AP"
        ''                    oJournalEntry.JournalEntries.Lines.UserFields.Fields.Item("U_AB_AP").Value = sBaseC
        ''                Case "CN"
        ''                    oJournalEntry.JournalEntries.Lines.UserFields.Fields.Item("U_AB_APCN").Value = sBaseC
        ''            End Select
        ''            oJournalEntry.JournalEntries.Lines.UserFields.Fields.Item("U_AB_JV").Value = sJC
        ''            If Not String.IsNullOrEmpty(sLOSC) Then
        ''                oJournalEntry.JournalEntries.Lines.CostingCode = sLOSC 'LOS
        ''            End If

        ''            If Not String.IsNullOrEmpty(sBUC) Then
        ''                oJournalEntry.JournalEntries.Lines.CostingCode2 = sBUC 'BU
        ''            End If

        ''            If Not String.IsNullOrEmpty(sOUC) Then
        ''                oJournalEntry.JournalEntries.Lines.CostingCode3 = sOUC 'OU
        ''            End If
        ''            If Not String.IsNullOrEmpty(sProjectC) Then
        ''                oJournalEntry.JournalEntries.Lines.ProjectCode = sProjectC 'Project
        ''            End If
        ''            oJournalEntry.JournalEntries.Lines.UserFields.Fields.Item("U_AB_NONPROJECT").Value = sNonprojectC  'OU_BU
        ''            oJournalEntry.JournalEntries.Lines.UserFields.Fields.Item("U_AB_REMARKS").Value = sLineRemarksC  'OU_BU
        ''            '' oJournalEntry.Lines.UserFields.Fields.Item("U_AB_NONPROJECT").Value = odr("OU_BU_Budget").ToString.Trim 'OU_BU
        ''            oJournalEntry.JournalEntries.Lines.Add()
        ''        Else
        ''            FCredit = False
        ''        End If
        ''        Console.WriteLine("Attempting to Add the Journal Entry ", sFuncName)

        ''        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add the Journal Entry", sFuncName)
        ''        ''  oCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
        ''        ''  oJournalEntry.SaveXML(System.Windows.Forms.Application.StartupPath & "\NONJVSource.xml")
        ''        ival = oJournalEntry.Add()

        ''        If ival <> 0 Then
        ''            IsError = True
        ''            oCompany.GetLastError(iErr, sErr)
        ''            Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
        ''            Console.WriteLine("Completed with ERROR ", sFuncName)
        ''            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
        ''            JournalEntry_Posting_NONJV_Source = RTN_ERROR
        ''            Throw New ArgumentException(sErr)
        ''        End If

        ''        Console.WriteLine("Completed with SUCCESS", sFuncName)
        ''        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS ", sFuncName)
        ''        oCompany.GetNewObjectCode(sRef)
        ''        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Journal Entry DocEntry  " & sRef, sFuncName)
        ''        JournalEntry_Posting_NONJV_Source = RTN_SUCCESS
        ''        sErrDesc = String.Empty

        ''    Catch ex As Exception
        ''        sErrDesc = ex.Message
        ''        Call WriteToLogFile(ex.Message, sFuncName)
        ''        Console.WriteLine("Completed with ERROR ", sFuncName)
        ''        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & ex.Message, sFuncName)
        ''        JournalEntry_Posting_NONJV_Source = RTN_ERROR
        ''        Exit Function
        ''    End Try

        ''End Function

        Public Function JournalEntry_Posting_NONJV_Source(ByVal oDVJour As DataView, ByRef oCompany As SAPbobsCOM.Company, ByVal sDate As String, ByRef sRef As String, ByRef sErrDesc As String) As Long

            Dim sFuncName As String = String.Empty
            Dim ival As Integer
            Dim IsError As Boolean
            Dim iErr As Integer = 0
            Dim sErr As String = String.Empty
            Dim sJV As String = String.Empty

            Dim oRset As SAPbobsCOM.Recordset = Nothing
            Dim sSQL As String = String.Empty
            Dim sPostingDate As String = String.Empty
            Dim oJournalEntry As SAPbobsCOM.JournalEntries = Nothing
            Dim bCredit, bDebit As Double
            Dim sBU As String = String.Empty
            Dim sOU As String = String.Empty
            Dim sLOS As String = String.Empty
            Dim sOUNAme As String = String.Empty
            Dim sLineRemarks As String = String.Empty
            Dim sNonproject As String = String.Empty
            Dim sProject As String = String.Empty
            Dim sBUC As String = String.Empty
            Dim sOUC As String = String.Empty
            Dim sLOSC As String = String.Empty
            Dim sOUNAmeC As String = String.Empty
            Dim sLineRemarksC As String = String.Empty
            Dim sNonprojectC As String = String.Empty
            Dim sProjectC As String = String.Empty
            Dim sGL As String = String.Empty
            Dim sCat As String = String.Empty
            Dim sBase As String = String.Empty
            Dim sCatC As String = String.Empty
            Dim sBaseC As String = String.Empty
            Dim sJC As String = String.Empty
            Dim dDate As Date
            Dim dCal As Double = 0
            Dim FCredit As Boolean = False
            Try
                sFuncName = "JournalEntry_Posting_NONJV_Source"
                Console.WriteLine("Starting Function ", sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)
                oRset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                sSQL = "SELECT T0.[AcctCode], T0.[AcctName] FROM OACT T0"
                oRset.DoQuery(sSQL)
                Dim odt As DataTable = Nothing
                odt = New DataTable()
                odt = ConvertRecordsetToDataTable(oRset, sErrDesc)
                Dim odv As DataView = Nothing
                odv = New DataView(odt)

                oJournalEntry = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
                oJournalEntry.ReferenceDate = DateTime.ParseExact(sDate, "yyyyMMdd", Nothing)
                oJournalEntry.DueDate = DateTime.ParseExact(sDate, "yyyyMMdd", Nothing)
                dDate = DateTime.ParseExact(sDate, "yyyyMMdd", Nothing)
                oJournalEntry.Memo = Left("Cost Allocation for the month of " & UCase(MonthName(Month(dDate))) & " - " & Year(dDate), 50)
                ''  oJournalEntry.Indicator = "CA"

                For Each odr As DataRowView In oDVJour
                    ''  If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug1(odr("GL_Code") & "  " & odr("GL_NameT") & "  " & odr("OU") & "  " & odr("Amount") & "  " & odr("TGL"), "JournalEntry_Posting_NONJV_Source Inside")
                    sGL = odr("TGL").ToString.Trim()
                    If odr("Amount") > 0 Then
                        ''  oJournalEntry.Lines.AccountCode = odr("TGL").ToString.Trim()
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug1("Debit", "")

                        bDebit += CDbl(odr("Amount").ToString.Trim)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug1("Debit Amount " & bDebit, "")
                        sCat = odr("Cat").ToString.Trim()
                        sBase = odr("Base").ToString.Trim()
                        sJV = odr("JV").ToString.Trim
                        sLOS = odr("LOS").ToString.Trim
                        sBU = odr("BU").ToString.Trim
                        sOU = odr("NewOU").ToString.Trim
                        sProject = odr("Project").ToString.Trim
                        sNonproject = odr("OU_BU_Budget").ToString.Trim
                        sLineRemarks = odr("Remarks").ToString.Trim.Replace("#$%", "'")

                        If sCat = "DEP" Then
                            odv.RowFilter = "AcctCode='" & P_sDEP_FAA & "'"
                            If odv.Count = 0 Then
                                Throw New ArgumentException(P_sDEP_FAA & "  - Account Code is missing in Source Entity " & oCompany.CompanyDB & " - " & oCompany.CompanyName)
                            End If
                            oJournalEntry.Lines.AccountCode = P_sDEP_FAA
                        Else
                            odv.RowFilter = "AcctCode='" & Left(odr("GL_NameT").ToString.Trim, Len(odr("GL_NameT").ToString.Trim) - 4) & "9999" & "'"
                            If odv.Count = 0 Then
                                Throw New ArgumentException(Left(odr("GL_NameT").ToString.Trim, Len(odr("GL_NameT").ToString.Trim) - 4) & "9999" & "  - Account Code is missing in Source Entity " & oCompany.CompanyDB & " - " & oCompany.CompanyName)
                            End If
                            oJournalEntry.Lines.AccountCode = Left(odr("GL_NameT").ToString.Trim, Len(odr("GL_NameT").ToString.Trim) - 4) & "9999"
                        End If
                        oJournalEntry.Lines.Credit = CDbl(odr("Amount").ToString.Trim)
                        Select Case odr("Cat").ToString.Trim()
                            Case "AP"
                                oJournalEntry.Lines.UserFields.Fields.Item("U_AB_AP").Value = odr("Base").ToString.Trim
                            Case "CN"
                                oJournalEntry.Lines.UserFields.Fields.Item("U_AB_APCN").Value = odr("Base").ToString.Trim
                        End Select
                        oJournalEntry.Lines.UserFields.Fields.Item("U_AB_JV").Value = odr("JV").ToString.Trim
                        If Not String.IsNullOrEmpty(odr("LOS").ToString.Trim) Then
                            oJournalEntry.Lines.CostingCode = odr("LOS").ToString.Trim 'LOS
                        End If

                        If Not String.IsNullOrEmpty(odr("BU").ToString.Trim) Then
                            oJournalEntry.Lines.CostingCode2 = odr("BU").ToString.Trim 'BU
                        End If

                        If Not String.IsNullOrEmpty(odr("NewOU").ToString.Trim) Then
                            oJournalEntry.Lines.CostingCode3 = odr("NewOU").ToString.Trim 'OU
                        End If
                        If Not String.IsNullOrEmpty(odr("Project").ToString.Trim) Then
                            oJournalEntry.Lines.ProjectCode = odr("Project").ToString.Trim 'Project
                        End If
                        oJournalEntry.Lines.UserFields.Fields.Item("U_AB_NONPROJECT").Value = odr("OU_BU_Budget").ToString.Trim 'OU_BU
                        oJournalEntry.Lines.UserFields.Fields.Item("U_AB_REMARKS").Value = odr("Remarks").ToString.Trim.Replace("#$%", "'") 'OU_BU
                        oJournalEntry.Lines.UserFields.Fields.Item("U_AB_OcrCode3").Value = odr("OcrCode3").ToString.Trim '
                        ''oJournalEntry.Lines.UserFields.Fields.Item("U_AB_NONPROJECT").Value = odr("OU_BU_Budget").ToString.Trim 'OU_BU
                        '' oJournalEntry.Lines.CostingCode = odr(18).ToString.Trim
                        '' oJournalEntry.Lines.CostingCode3 = odr(10).ToString.Trim
                    Else
                        ''  If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug1("Credit", "")

                        bCredit += Math.Abs(CDbl(odr("Amount").ToString.Trim))
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug1("Credit Amount " & bCredit, "")
                        sCatC = odr("Cat").ToString.Trim()
                        sBaseC = odr("Base").ToString.Trim
                        sJC = odr("JV").ToString.Trim
                        sLOSC = odr("LOS").ToString.Trim
                        sBUC = odr("BU").ToString.Trim
                        sOUC = odr("NewOU").ToString.Trim
                        sProjectC = odr("Project").ToString.Trim
                        sNonprojectC = odr("OU_BU_Budget").ToString.Trim
                        sLineRemarksC = odr("Remarks").ToString.Trim.Replace("#$%", "'")

                        If sCat = "DEP" Then
                            odv.RowFilter = "AcctCode='" & P_sDEP_FAA & "'"
                            If odv.Count = 0 Then
                                Throw New ArgumentException(P_sDEP_FAA & "  - Account Code is missing in Source Entity " & oCompany.CompanyDB & " - " & oCompany.CompanyName)
                            End If
                            oJournalEntry.Lines.AccountCode = P_sDEP_FAA
                        Else
                            odv.RowFilter = "AcctCode='" & Left(odr("GL_NameT").ToString.Trim, Len(odr("GL_NameT").ToString.Trim) - 4) & "9999" & "'"
                            If odv.Count = 0 Then
                                Throw New ArgumentException(Left(odr("GL_NameT").ToString.Trim, Len(odr("GL_NameT").ToString.Trim) - 4) & "9999" & "  - Account Code is missing in Source Entity " & oCompany.CompanyDB & " - " & oCompany.CompanyName)
                            End If
                            oJournalEntry.Lines.AccountCode = Left(odr("GL_NameT").ToString.Trim, Len(odr("GL_NameT").ToString.Trim) - 4) & "9999"
                        End If
                        '' oJournalEntry.Lines.AccountCode = Left(odr("GL_NameT").ToString.Trim, Len(odr("GL_NameT").ToString.Trim) - 4) & "9999"
                        oJournalEntry.Lines.Debit = Math.Abs(CDbl(odr("Amount").ToString.Trim))
                        Select Case odr("Cat").ToString.Trim()
                            Case "AP"
                                oJournalEntry.Lines.UserFields.Fields.Item("U_AB_AP").Value = odr("Base").ToString.Trim
                            Case "CN"
                                oJournalEntry.Lines.UserFields.Fields.Item("U_AB_APCN").Value = odr("Base").ToString.Trim
                        End Select
                        oJournalEntry.Lines.UserFields.Fields.Item("U_AB_JV").Value = odr("JV").ToString.Trim
                        If Not String.IsNullOrEmpty(odr("LOS").ToString.Trim) Then
                            oJournalEntry.Lines.CostingCode = odr("LOS").ToString.Trim 'LOS
                        End If

                        If Not String.IsNullOrEmpty(odr("BU").ToString.Trim) Then
                            oJournalEntry.Lines.CostingCode2 = odr("BU").ToString.Trim 'BU
                        End If

                        If Not String.IsNullOrEmpty(odr("NewOU").ToString.Trim) Then
                            oJournalEntry.Lines.CostingCode3 = odr("NewOU").ToString.Trim 'OU
                        End If
                        If Not String.IsNullOrEmpty(odr("Project").ToString.Trim) Then
                            oJournalEntry.Lines.ProjectCode = odr("Project").ToString.Trim 'Project
                        End If
                        oJournalEntry.Lines.UserFields.Fields.Item("U_AB_NONPROJECT").Value = odr("OU_BU_Budget").ToString.Trim 'OU_BU
                        oJournalEntry.Lines.UserFields.Fields.Item("U_AB_REMARKS").Value = odr("Remarks").ToString.Trim.Replace("#$%", "'") 'OU_BU
                        oJournalEntry.Lines.UserFields.Fields.Item("U_AB_OcrCode3").Value = odr("OcrCode3").ToString.Trim '
                        ''  oJournalEntry.Lines.UserFields.Fields.Item("U_AB_NONPROJECT").Value = odr("OU_BU_Budget").ToString.Trim 'OU_BU
                        '' oJournalEntry.Lines.CostingCode = odr(18).ToString.Trim
                        '' oJournalEntry.Lines.CostingCode3 = odr(10).ToString.Trim
                    End If
                    oJournalEntry.Lines.Add()
                Next

                If bDebit > 0 Then
                    odv.RowFilter = "AcctCode='" & sGL & "'"
                    If odv.Count = 0 Then
                        Throw New ArgumentException(sGL & "  - Account Code is missing in Source Entity " & oCompany.CompanyDB & " - " & oCompany.CompanyName)
                    End If

                    oJournalEntry.Lines.AccountCode = sGL
                    If bCredit > 0 Then
                        dCal = bDebit - bCredit
                        If dCal > 0 Then
                            oJournalEntry.Lines.Debit = dCal
                        Else
                            oJournalEntry.Lines.Credit = Math.Abs(dCal)
                        End If
                        FCredit = True
                    Else
                        oJournalEntry.Lines.Debit = bDebit
                    End If
                    Select Case sCat
                        Case "AP"
                            oJournalEntry.Lines.UserFields.Fields.Item("U_AB_AP").Value = sBase
                        Case "CN"
                            oJournalEntry.Lines.UserFields.Fields.Item("U_AB_APCN").Value = sBase
                    End Select
                    oJournalEntry.Lines.UserFields.Fields.Item("U_AB_JV").Value = sJV
                    If Not String.IsNullOrEmpty(sLOS) Then
                        oJournalEntry.Lines.CostingCode = sLOS 'LOS
                    End If

                    If Not String.IsNullOrEmpty(sBU) Then
                        oJournalEntry.Lines.CostingCode2 = sBU 'BU
                    End If

                    If Not String.IsNullOrEmpty(sOU) Then
                        oJournalEntry.Lines.CostingCode3 = sOU 'OU
                    End If
                    If Not String.IsNullOrEmpty(sProject) Then
                        oJournalEntry.Lines.ProjectCode = sProject  'Project
                    End If
                    oJournalEntry.Lines.UserFields.Fields.Item("U_AB_NONPROJECT").Value = sNonproject  'OU_BU
                    oJournalEntry.Lines.UserFields.Fields.Item("U_AB_REMARKS").Value = sLineRemarks  'OU_BU
                    ''oJournalEntry.Lines.UserFields.Fields.Item("U_AB_NONPROJECT").Value = odr("OU_BU_Budget").ToString.Trim 'OU_BU
                    oJournalEntry.Lines.Add()
                End If

                If bCredit > 0 And FCredit = False Then
                    odv.RowFilter = "AcctCode='" & sGL & "'"
                    If odv.Count = 0 Then
                        Throw New ArgumentException(sGL & "  - Account Code is missing in Source Entity " & oCompany.CompanyDB & " - " & oCompany.CompanyName)
                    End If
                    oJournalEntry.Lines.AccountCode = sGL
                    oJournalEntry.Lines.Credit = bCredit
                    Select Case sCatC
                        Case "AP"
                            oJournalEntry.Lines.UserFields.Fields.Item("U_AB_AP").Value = sBaseC
                        Case "CN"
                            oJournalEntry.Lines.UserFields.Fields.Item("U_AB_APCN").Value = sBaseC
                    End Select
                    oJournalEntry.Lines.UserFields.Fields.Item("U_AB_JV").Value = sJC
                    If Not String.IsNullOrEmpty(sLOSC) Then
                        oJournalEntry.Lines.CostingCode = sLOSC 'LOS
                    End If

                    If Not String.IsNullOrEmpty(sBUC) Then
                        oJournalEntry.Lines.CostingCode2 = sBUC 'BU
                    End If

                    If Not String.IsNullOrEmpty(sOUC) Then
                        oJournalEntry.Lines.CostingCode3 = sOUC 'OU
                    End If
                    If Not String.IsNullOrEmpty(sProjectC) Then
                        oJournalEntry.Lines.ProjectCode = sProjectC 'Project
                    End If
                    oJournalEntry.Lines.UserFields.Fields.Item("U_AB_NONPROJECT").Value = sNonprojectC  'OU_BU
                    oJournalEntry.Lines.UserFields.Fields.Item("U_AB_REMARKS").Value = sLineRemarksC  'OU_BU
                    '' oJournalEntry.Lines.UserFields.Fields.Item("U_AB_NONPROJECT").Value = odr("OU_BU_Budget").ToString.Trim 'OU_BU
                    oJournalEntry.Lines.Add()
                Else
                    FCredit = False
                End If
                Console.WriteLine("Attempting to Add the Journal Entry ", sFuncName)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add the Journal Entry", sFuncName)
                oCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
                oJournalEntry.SaveXML(System.Windows.Forms.Application.StartupPath & "\NONJVSource.xml")
                ival = oJournalEntry.Add()

                If ival <> 0 Then
                    IsError = True
                    oCompany.GetLastError(iErr, sErr)
                    Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
                    Console.WriteLine("Completed with ERROR ", sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
                    JournalEntry_Posting_NONJV_Source = RTN_ERROR
                    Throw New ArgumentException(sErr)
                End If

                Console.WriteLine("Completed with SUCCESS", sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS ", sFuncName)
                oCompany.GetNewObjectCode(sRef)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Journal Entry DocEntry  " & sRef, sFuncName)
                JournalEntry_Posting_NONJV_Source = RTN_SUCCESS
                sErrDesc = String.Empty

            Catch ex As Exception
                sErrDesc = ex.Message
                Call WriteToLogFile(ex.Message, sFuncName)
                Console.WriteLine("Completed with ERROR ", sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & ex.Message, sFuncName)
                JournalEntry_Posting_NONJV_Source = RTN_ERROR
                Exit Function
            End Try

        End Function

        Public Function JournalEntry_Posting_NONJV_Target(ByVal oDVJour As DataView, ByRef oCompany As SAPbobsCOM.Company, ByVal sDate As String, ByVal sRef As String, ByRef sErrDesc As String) As Long

            Dim sFuncName As String = String.Empty
            Dim ival As Integer
            Dim IsError As Boolean
            Dim iErr As Integer = 0
            Dim sErr As String = String.Empty
            Dim sJV As String = String.Empty

            Dim bDebit As Double = 0.0
            Dim bCredit As Double = 0.0

            Dim sBU As String = String.Empty
            Dim sOU As String = String.Empty
            Dim sLOS As String = String.Empty
            Dim sOUNAme As String = String.Empty
            Dim sLineRemarks As String = String.Empty
            Dim sNonproject As String = String.Empty
            Dim sProject As String = String.Empty
            Dim sBUC As String = String.Empty
            Dim sOUC As String = String.Empty
            Dim sLOSC As String = String.Empty
            Dim sOUNAmeC As String = String.Empty
            Dim sLineRemarksC As String = String.Empty
            Dim sNonprojectC As String = String.Empty
            Dim sProjectC As String = String.Empty
            Dim dDate As Date

            Dim oRset As SAPbobsCOM.Recordset = Nothing
            Dim sSQL As String = String.Empty
            Dim sPostingDate As String = String.Empty
            Dim oJournalEntry As SAPbobsCOM.JournalVouchers = Nothing
            Dim oDTJE As DataTable = Nothing
            Dim sAccCode As String = Nothing
            Dim FCredit As Boolean = False
            Dim dCal As Double = 0

            Try
                sFuncName = "JournalEntry_Posting_NONJV_Target"
                Console.WriteLine("Starting Function ", sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)
                oRset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                sSQL = "SELECT T0.[AcctCode], T0.[AcctName] FROM OACT T0"
                oRset.DoQuery(sSQL)
                Dim odt As DataTable = Nothing
                odt = New DataTable()
                odt = ConvertRecordsetToDataTable(oRset, sErrDesc)
                Dim odv As DataView = Nothing
                odv = New DataView(odt)

                oJournalEntry = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalVouchers)
                oDTJE = New DataTable()

                oDTJE.Columns.Add("AccountCode", GetType(String))
                oDTJE.Columns.Add("Debit", GetType(Double))
                oDTJE.Columns.Add("Credit", GetType(Double))
                oDTJE.Columns.Add("CostingCode", GetType(String))
                oDTJE.Columns.Add("CostingCode2", GetType(String))
                oDTJE.Columns.Add("CostingCode3", GetType(String))
                oDTJE.Columns.Add("CostingCode4", GetType(String))
                oDTJE.Columns.Add("LineMemo", GetType(String))
                dDate = DateTime.ParseExact(sDate, "yyyyMMdd", Nothing)

                For Each odr As DataRowView In oDVJour
                    If odr("Amount") > 0 Then
                        If odr("Cat").ToString().Trim() = "DEP" Then
                            sAccCode = P_sDEP_FAA
                        Else
                            sAccCode = odr("GL_NameT").ToString.Trim
                        End If

                        oDTJE.Rows.Add(sAccCode, CDbl(odr("Amount").ToString.Trim), 0, odr("LOS").ToString.Trim, odr("BU").ToString.Trim, odr("NewOU").ToString.Trim, odr("Project").ToString.Trim, odr("Remarks").ToString.Trim.Replace("#$%", "'"))
                        bCredit += CDbl(odr("Amount").ToString.Trim)
                        sLOSC = odr("LOS").ToString.Trim
                        sBUC = odr("BU").ToString.Trim
                        sOUC = odr("NewOU").ToString.Trim
                        sProjectC = odr("Project").ToString.Trim
                        sLineRemarksC = odr("Remarks").ToString.Trim.Replace("#$%", "'")
                    Else
                        If odr("Cat").ToString().Trim() = "DEP" Then
                          
                            sAccCode = P_sDEP_FAA
                        Else
                            
                            sAccCode = odr("GL_NameT").ToString.Trim
                        End If
                        oDTJE.Rows.Add(sAccCode, 0, Math.Abs(CDbl(odr("Amount").ToString.Trim)), odr("LOS").ToString.Trim, odr("BU").ToString.Trim, odr("NewOU").ToString.Trim, odr("Project").ToString.Trim, odr("Remarks").ToString.Trim.Replace("#$%", "'"))
                        'oJournalEntry.JournalEntries.Lines.AccountCode = odr("GL_NameT").ToString.Trim
                        bDebit += Math.Abs(CDbl(odr("Amount").ToString.Trim))
                        sLOS = odr("LOS").ToString.Trim 'LOS
                        sBU = odr("BU").ToString.Trim 'BU
                        sOU = odr("NewOU").ToString.Trim 'OU
                        sProject = odr("Project").ToString.Trim 'Project
                        sLineRemarks = odr("Remarks").ToString.Trim.Replace("#$%", "'") 'Project
                    End If

                Next

                If bDebit > 0 Then
                    If bCredit > 0 Then
                        dCal = bDebit - bCredit
                        If dCal > 0 Then
                            oDTJE.Rows.Add(P_sNonJV_Credit, dCal, 0, sLOS, sBU, sOU, sProject, Left("Cost Allocation for the month of " & UCase(MonthName(Month(dDate))) & " - " & Year(dDate), 50))
                        Else
                            oDTJE.Rows.Add(P_sNonJV_Credit, 0, Math.Abs(dCal), sLOS, sBU, sOU, sProject, Left("Cost Allocation for the month of " & UCase(MonthName(Month(dDate))) & " - " & Year(dDate), 50))
                        End If

                        '' oDTJE.Rows.Add(P_sNonJV_Credit, bDebit, bCredit, sLOS, sBU, sOU, sProject, Left("Cost Allocation for the month of " & UCase(MonthName(Month(dDate))) & " - " & Year(dDate), 50))
                        FCredit = True
                    Else
                        oDTJE.Rows.Add(P_sNonJV_Credit, bDebit, 0, sLOS, sBU, sOU, sProject, sLineRemarks)
                    End If
                End If
                If bCredit > 0 And FCredit = False Then
                    oDTJE.Rows.Add(P_sNonJV_Credit, 0, bCredit, sLOSC, sBUC, sOUC, sProjectC, sLineRemarksC)
                Else
                    FCredit = False
                End If

                For Each oddr As DataRow In oDTJE.Rows
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug1(oddr("AccountCode") & "  " & oddr("Debit") & "  " & oddr("Credit"), "JournalEntry_Posting_NONJV_Target_Inside")
                Next

                oJournalEntry.JournalEntries.ReferenceDate = DateTime.ParseExact(sDate, "yyyyMMdd", Nothing)
                oJournalEntry.JournalEntries.DueDate = DateTime.ParseExact(sDate, "yyyyMMdd", Nothing)
                oJournalEntry.JournalEntries.Memo = Left("Cost Allocation for the month of " & UCase(MonthName(Month(dDate))) & " - " & Year(dDate), 50)
                oJournalEntry.JournalEntries.Reference3 = sRef



                For Each odr As DataRow In oDTJE.Rows
                    odv.RowFilter = "AcctCode='" & odr("AccountCode").ToString.Trim & "'"
                    If odv.Count = 0 Then
                        Throw New ArgumentException(odr("AccountCode").ToString.Trim & "  - Account Code is missing in Traget Entity " & oCompany.CompanyDB & " - " & oCompany.CompanyName)
                    End If
                    oJournalEntry.JournalEntries.Lines.AccountCode = odr("AccountCode").ToString.Trim
                    oJournalEntry.JournalEntries.Lines.Debit = CDbl(odr("Debit").ToString.Trim)
                    oJournalEntry.JournalEntries.Lines.Credit = CDbl(odr("Credit").ToString.Trim)
                    If Not String.IsNullOrEmpty(odr("CostingCode").ToString.Trim) Then
                        oJournalEntry.JournalEntries.Lines.CostingCode = odr("CostingCode").ToString.Trim 'LOS
                    End If
                    If Not String.IsNullOrEmpty(odr("CostingCode2").ToString.Trim) Then
                        oJournalEntry.JournalEntries.Lines.CostingCode2 = odr("CostingCode2").ToString.Trim 'BU
                    End If
                    If Not String.IsNullOrEmpty(odr("CostingCode3").ToString.Trim) Then
                        oJournalEntry.JournalEntries.Lines.CostingCode3 = odr("CostingCode3").ToString.Trim 'OU
                    End If
                    If Not String.IsNullOrEmpty(odr("CostingCode4").ToString.Trim) Then
                        oJournalEntry.JournalEntries.Lines.ProjectCode = odr("CostingCode4").ToString.Trim 'Project
                    End If
                    oJournalEntry.JournalEntries.Lines.LineMemo = odr("LineMemo").ToString.Trim
                    oJournalEntry.JournalEntries.Lines.Add()
                Next



                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add the Journal Entry", sFuncName)
                ival = oJournalEntry.Add()

                If ival <> 0 Then
                    IsError = True
                    oCompany.GetLastError(iErr, sErr)
                    Call WriteToLogFile("Completed with ERROR ---" & sErr, sFuncName)
                    Console.WriteLine("Completed with ERROR ", sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & sErr, sFuncName)
                    JournalEntry_Posting_NONJV_Target = RTN_ERROR
                    Throw New ArgumentException(sErr)
                End If

                Console.WriteLine("Completed with SUCCESS", sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS ", sFuncName)
                oCompany.GetNewObjectCode(sJV)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Journal Entry DocEntry  " & sJV, sFuncName)
                sErrDesc = String.Empty
                JournalEntry_Posting_NONJV_Target = RTN_SUCCESS

            Catch ex As Exception
                sErrDesc = ex.Message

                Call WriteToLogFile(ex.Message, sFuncName)
                Console.WriteLine("Completed with ERROR ", sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & ex.Message, sFuncName)
                JournalEntry_Posting_NONJV_Target = RTN_ERROR
                Exit Function
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
                Dim ss = p_oDICompany.Server
                Dim dd = p_oDICompany.DbServerType
                Dim aa = p_oDICompany.LicenseServer

                oCompany.Server = p_oDICompany.Server
                oCompany.DbServerType = p_oDICompany.DbServerType
                oCompany.LicenseServer = p_oDICompany.LicenseServer
                oCompany.CompanyDB = sEntity
                oCompany.UserName = sUsername
                oCompany.Password = sPassword
                oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English
                oCompany.UseTrusted = False
                oCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
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

        Public Sub PrintCalling(ByVal oDS As DataSet)
            Try
                sFuncName = "PrintCalling()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

                Dim cryRpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument
                Dim crtableLogoninfos As New CrystalDecisions.Shared.TableLogOnInfos
                Dim crtableLogoninfo As New CrystalDecisions.Shared.TableLogOnInfo
                Dim crConnectionInfo As New CrystalDecisions.Shared.ConnectionInfo
                Dim CrTables As CrystalDecisions.CrystalReports.Engine.Tables
                Dim CrTable As CrystalDecisions.CrystalReports.Engine.Table
                Dim sPath As String
                sPath = IO.Directory.GetParent(Application.StartupPath).ToString

                'MsgBox(System.Windows.Forms.Application.StartupPath.ToString)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Loading Layout ", sFuncName)
                Dim dd = System.Windows.Forms.Application.StartupPath.ToString & "\" & "JournalEntryInfo.rpt"
                cryRpt.Load(System.Windows.Forms.Application.StartupPath.ToString & "\" & "JournalEntryInfo.rpt")
                'cryRpt.Load("PUT CRYSTAL REPORT PATH HERE\CrystalReport1.rpt")

                Dim crParameterFieldDefinitions As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinitions
                Dim crParameterFieldDefinition As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinition
                Dim crParameterValues As New CrystalDecisions.Shared.ParameterValues
                Dim crParameterDiscreteValue As New CrystalDecisions.Shared.ParameterDiscreteValue

                Dim RptFrm As Viewer
                RptFrm = New Viewer
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Assigning DS to Layout ", sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("DS Count " & oDS.Tables("JE").Rows.Count, sFuncName)
                cryRpt.SetDataSource(oDS.Tables("JE"))
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("1 ", sFuncName)
                RptFrm.CrystalReportViewer1.ReportSource = cryRpt
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("2 ", sFuncName)
                RptFrm.CrystalReportViewer1.Refresh()
                RptFrm.Text = "Journal Entry Information "
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Printing Layout", sFuncName)
                Dim oPS As New System.Drawing.Printing.PrinterSettings
                Dim PrinterName As String = String.Empty
                PrinterName = oPS.PrinterName
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Default Printer name " & PrinterName, sFuncName)
                cryRpt.PrintOptions.PrinterName = PrinterName
                cryRpt.PrintToPrinter(1, False, 1, 100)
                ''  RptFrm.TopMost = True

                ''  RptFrm.Activate()
                '' RptFrm.ShowDialog()
                System.Threading.Thread.Sleep(100)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

            Catch ex As Exception
                p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End Try

        End Sub

        Public Sub PrintCalling_Summary(ByVal oDS As DataSet)
            Try
                sFuncName = "PrintCalling()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

                Dim cryRpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument
                Dim crtableLogoninfos As New CrystalDecisions.Shared.TableLogOnInfos
                Dim crtableLogoninfo As New CrystalDecisions.Shared.TableLogOnInfo
                Dim crConnectionInfo As New CrystalDecisions.Shared.ConnectionInfo
                Dim CrTables As CrystalDecisions.CrystalReports.Engine.Tables
                Dim CrTable As CrystalDecisions.CrystalReports.Engine.Table
                Dim sPath As String
                sPath = IO.Directory.GetParent(Application.StartupPath).ToString

                'MsgBox(System.Windows.Forms.Application.StartupPath.ToString)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Loading Layout ", sFuncName)
                ''  Dim dd = System.Windows.Forms.Application.StartupPath.ToString & "\" & "JournalEntryInfo.rpt"
                cryRpt.Load(System.Windows.Forms.Application.StartupPath.ToString & "\" & "Summaryreport.rpt")
                'cryRpt.Load("PUT CRYSTAL REPORT PATH HERE\CrystalReport1.rpt")

                Dim crParameterFieldDefinitions As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinitions
                Dim crParameterFieldDefinition As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinition
                Dim crParameterValues As New CrystalDecisions.Shared.ParameterValues
                Dim crParameterDiscreteValue As New CrystalDecisions.Shared.ParameterDiscreteValue

                Dim RptFrm As Viewer
                RptFrm = New Viewer

                ''For Each odr As DataRow In oDS.Tables(0).Rows
                ''    crParameterDiscreteValue.Value = 0
                ''    crParameterFieldDefinitions = _
                ''cryRpt.DataDefinition.ParameterFields
                ''    crParameterFieldDefinition = _
                ''crParameterFieldDefinitions.Item("")
                ''    crParameterValues = crParameterFieldDefinition.CurrentValues

                ''    crParameterValues.Clear()
                ''    crParameterValues.Add(crParameterDiscreteValue)
                ''    crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)
                ''Next



                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Assinging DS to Layout", sFuncName)
                cryRpt.SetDataSource(oDS.Tables("Summary"))
                RptFrm.CrystalReportViewer1.ReportSource = cryRpt
                RptFrm.CrystalReportViewer1.Refresh()
                RptFrm.Text = "Summary Report "

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Printing", sFuncName)
                Dim oPS As New System.Drawing.Printing.PrinterSettings
                Dim PrinterName As String = String.Empty
                PrinterName = oPS.PrinterName
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Default Printer name " & PrinterName, sFuncName)
                cryRpt.PrintOptions.PrinterName = PrinterName
                cryRpt.PrintToPrinter(1, False, 1, 100)
                ''  RptFrm.TopMost = True

                ''  RptFrm.Activate()
                '' RptFrm.ShowDialog()
                System.Threading.Thread.Sleep(100)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

            Catch ex As Exception
                p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End Try

        End Sub

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

        Public Function CostAllocation_Validation(ByRef oForm As SAPbouiCOM.Form, ByRef sErrDesc As String) As Long
            Try

                sFuncName = "CostAllocation_Validation()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

                If String.IsNullOrEmpty(oForm.Items.Item("Item_13").Specific.value) Then
                    sErrDesc = "Month From couldn`t be blank ....! "
                    Return RTN_ERROR
                End If
                If String.IsNullOrEmpty(oForm.Items.Item("Item_14").Specific.value) Then
                    sErrDesc = "Month To couldn`t be blank ....! "
                    Return RTN_ERROR
                End If
                If String.IsNullOrEmpty(p_Dimensionrules) Then
                    sErrDesc = "Distribution rule couldn`t be blank ....!  "
                    Return RTN_ERROR
                End If
                If String.IsNullOrEmpty(oForm.Items.Item("Item_10").Specific.string) Then
                    sErrDesc = "GL Account From couldn`t be blank ....!"
                    Return RTN_ERROR
                End If
                If String.IsNullOrEmpty(oForm.Items.Item("Item_12").Specific.string) Then
                    sErrDesc = "GL Account To couldn`t be blank ....!"
                    Return RTN_ERROR
                End If
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
                Return RTN_SUCCESS
            Catch ex As Exception
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                sErrDesc = ex.Message
                Return RTN_ERROR
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

        Public Function ConvertRecordsetToDataTable(ByVal SAPRecordset As SAPbobsCOM.Recordset, ByRef sErrDesc As String) As DataTable

            '\ This function will take an SAP recordset from the SAPbobsCOM library and convert it to a more
            '\ easily used ADO.NET datatable which can be used for data binding much easier.
            Dim NewCol As DataColumn
            Dim NewRow As DataRow
            Dim ColCount As Integer
            Dim dAmount As Decimal = 0.0
            Dim dtTable As DataTable = Nothing

            sFuncName = "ConvertRecordsetToDataTable()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Fuction ", sFuncName)

            Try
                dtTable = New DataTable()

                For ColCount = 0 To SAPRecordset.Fields.Count - 1
                    NewCol = New DataColumn(SAPRecordset.Fields.Item(ColCount).Name)
                    '' If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Column Creation  ", sFuncName)
                    Try
                        dtTable.Columns.Add(NewCol)
                    Catch ex As Exception
                    End Try
                Next

                Do Until SAPRecordset.EoF
                    NewRow = dtTable.NewRow
                    'populate each column in the row we're creating
                    For ColCount = 0 To SAPRecordset.Fields.Count - 1
                        '' NewRow.Item(SAPRecordset.Fields.Item(ColCount).Name) = Convert.ToDecimal(SAPRecordset.Fields.Item(ColCount).Value)
                        NewRow.Item(SAPRecordset.Fields.Item(ColCount).Name) = SAPRecordset.Fields.Item(ColCount).Value
                    Next
                    'Add the row to the datatable
                    dtTable.Rows.Add(NewRow)
                    SAPRecordset.MoveNext()
                Loop

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS ", sFuncName)
                Return dtTable
                sErrDesc = String.Empty

            Catch ex As Exception
                sErrDesc = ex.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                p_oSBOApplication.StatusBar.SetText(ex.ToString & Chr(10) & "Error converting SAP Recordset to DataTable", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return Nothing
            End Try


        End Function

        Public Function Write_TextFileError(ByVal oDT_FinalResult As DataTable, ByVal sPAth As String, ByRef sErrDesc As String) As Long
            Try
                Dim sFuncName As String = String.Empty
                Dim irow As Integer
                Dim sFileName As String = "\ValidationError.txt"
                Dim sbuffer As String = String.Empty
                Dim sline As String = "="
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
                sw.WriteLine("Validation Error ")
                sw.WriteLine("      ")
                sw.WriteLine("      ")
                sw.WriteLine("OU Code        " & Space(10) & "Error Msg")
                sw.WriteLine(sline.PadRight(100, "="c))
                sw.WriteLine("      ")

                For imjs = 0 To oDT_FinalResult.Rows.Count - 1
                    sw.WriteLine(oDT_FinalResult.Rows(imjs).Item("OU").ToString & Space(15) & oDT_FinalResult.Rows(imjs).Item("Error").ToString)
                Next imjs
                sw.Close()
                Process.Start(sPAth & sFileName)

                Write_TextFileError = RTN_SUCCESS
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With SUCCESS ", sFuncName)

            Catch ex As Exception
                Write_TextFileError = RTN_ERROR
                sErrDesc = ex.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
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

        Public Sub ExportToExcel(ByVal dtTemp As DataTable, ByVal filepath As String)

            Dim strFileName As String = filepath
            Dim _excel As New Excel.Application
            Dim wBook As Excel.Workbook
            Dim wSheet As Excel.Worksheet

            Try

                Try
                    If File.Exists(filepath) Then
                        File.Delete(filepath)
                    End If
                Catch ex As Exception

                End Try

                wBook = _excel.Workbooks.Add()
                wSheet = wBook.ActiveSheet()

                Dim dt As System.Data.DataTable = dtTemp
                Dim dc As System.Data.DataColumn
                Dim dr As System.Data.DataRow
                Dim colIndex As Integer = 0
                Dim rowIndex As Integer = 0

                For Each dc In dt.Columns
                    colIndex = colIndex + 1
                    wSheet.Cells(1, colIndex) = dc.ColumnName
                Next

                For Each dr In dt.Rows
                    rowIndex = rowIndex + 1
                    colIndex = 0
                    For Each dc In dt.Columns
                        colIndex = colIndex + 1
                        wSheet.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName)
                    Next
                Next
                wSheet.Columns.AutoFit()
                wBook.SaveAs(strFileName)
                Process.Start(strFileName)

            Catch ex As Exception
            Finally
                ReleaseObject(wSheet)
                wBook.Close(False)
                ReleaseObject(wBook)
                _excel.Quit()
                ReleaseObject(_excel)
            End Try
        End Sub
        Private Sub ReleaseObject(ByVal o As Object)
            Try
                While (System.Runtime.InteropServices.Marshal.ReleaseComObject(o) > 0)
                End While
            Catch
            Finally
                o = Nothing
            End Try
        End Sub

    End Module
End Namespace



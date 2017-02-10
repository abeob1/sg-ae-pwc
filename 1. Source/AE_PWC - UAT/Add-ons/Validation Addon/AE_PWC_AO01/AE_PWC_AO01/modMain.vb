Option Explicit On



Module modMain


    Public p_iDebugMode As Int16
    Public p_iErrDispMethod As Int16
    Public p_iDeleteDebugLog As Int16

    Public Const RTN_SUCCESS As Int16 = 1
    Public Const RTN_ERROR As Int16 = 0

    Public Const DEBUG_ON As Int16 = 1
    Public Const DEBUG_OFF As Int16 = 0

    Public Const ERR_DISPLAY_STATUS As Int16 = 1
    Public Const ERR_DISPLAY_DIALOGUE As Int16 = 2

    Public Structure CompanyDefault
        'Public sSMTPServer As String
        'Public sSMTPPort As String
        'Public sEmailFrom As String
        'Public sSMTPUser As String
        'Public sSMTPPassword As String
        Public sAuthorization As String
        Public sApprover As String

    End Structure

    Public p_oApps As SAPbouiCOM.SboGuiApi
    Public p_oEventHandler As clsEventHandler
    Public WithEvents p_oSBOApplication As SAPbouiCOM.Application
    Public p_oDICompany As SAPbobsCOM.Company
    Public p_oUICompany As SAPbouiCOM.Company
    Public p_oCompDef As CompanyDefault
    Public sFuncName As String
    Public sErrDesc As String
    Public Approval As Boolean = False
    Public P_JEReverse As Boolean = False
    Public P_JEReversetmp As Boolean = False


    Public p_sSelectedFilepath As String = String.Empty
    Public p_sSelectedFileName As String = String.Empty
    Public p_sRefNuber(100, 4) As String
    'Public p_iArrayCount As Integer = 0
    'Public p_iArrayAcctCount As Integer = 0
    'Public p_iArrayAcctActiveCount As Integer = 0
    'Public p_sAccountCodes(100) As String
    'Public p_sAccountCodes_ActiveAccount(100) As String
    Public p_FormTypecount As Integer = 0
    Public p_BPTypecount As Integer = 0
    Public p_sEmailAddress As String = String.Empty
    Public p_sHoldingEntity As String = String.Empty

    Public p_oDTPOMatrixs As New DataTable
    Public p_oDTConsBudget As DataTable
    Public p_oDTEmailAddress As DataTable = Nothing
    Public p_sAstatus As String = String.Empty
    Public p_sDocType As String = String.Empty
    Public p_POApprovalCode As String = String.Empty
    Public p_PRApprovalCode As String = String.Empty


    

    Sub main(ByVal Args() As String)

        Try

            sFuncName = "Main()"
            p_sHoldingEntity = "PWCL"

            p_iDebugMode = DEBUG_ON
            p_iErrDispMethod = ERR_DISPLAY_STATUS

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Addon startup function", sFuncName)
            p_oApps = New SAPbouiCOM.SboGuiApi
            'p_oApps.Connect(Args(0))

            Dim sconn As String = Environment.GetCommandLineArgs.GetValue(1)
            p_oApps.Connect(sconn)


            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Initializing public SBO Application object", sFuncName)
            p_oSBOApplication = p_oApps.GetApplication

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Retriving SBO application company handle", sFuncName)
            p_oUICompany = p_oSBOApplication.Company


            p_oDICompany = New SAPbobsCOM.Company
            If Not p_oDICompany.Connected Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectDICompSSO()", sFuncName)
                If ConnectDICompSSO(p_oDICompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddMenus Functions", sFuncName)
            'Call AddMenuItems()
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("AddMenus Functions Completed Successfully.", sFuncName)

            '' Call WriteToLogFile_Debug("Calling DisplayStatus()", sFuncName)
            'Call DisplayStatus(Nothing, "Addon starting.....please wait....", sErrDesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating Event handler class", sFuncName)
            p_oEventHandler = New clsEventHandler(p_oSBOApplication, p_oDICompany)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling SetApplication Function", sFuncName)
            ' Call p_oEventHandler.SetApplication(sErrDesc)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling GetSystemIntializeInfo()", sFuncName)
            If GetSystemIntializeInfo(p_oCompDef, p_oDICompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Addon started successfully", "Main()")

            p_oDTPOMatrixs.Columns.Add("Sno", GetType(Integer))
            p_oDTPOMatrixs.Columns.Add("GLAccount", GetType(String))
            p_oDTPOMatrixs.Columns.Add("LineAmount", GetType(Decimal))
            p_oDTPOMatrixs.Columns.Add("OU", GetType(String))
            p_oDTPOMatrixs.Columns.Add("BU", GetType(String))
            p_oDTPOMatrixs.Columns.Add("Project", GetType(String))
            p_oDTPOMatrixs.Columns.Add("Cat", GetType(String))
            p_oDTPOMatrixs.Columns.Add("UpdateAmount", GetType(Decimal))
            p_oDTPOMatrixs.Columns.Add("DocEntry", GetType(String))

            Call WriteToLogFile_Debug("Calling EndStatus()", sFuncName)
            ' Call EndStatus(sErrDesc)
            p_oSBOApplication.StatusBar.SetText("Addon Started Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            System.Windows.Forms.Application.Run()



        Catch exp As Exception
            Call WriteToLogFile(exp.Message, "Main()")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", "Main()")
        Finally
        End Try
    End Sub

End Module






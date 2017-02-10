Module modMain

#Region "Varibles Declarations"

    'Company Default Structure

    Public Structure CompanyDefault

        Public sServer As String
        Public sLicenseServer As String
        Public sDBName As String
        Public sServerType As String
        Public iServerLanguage As Integer
        Public sSAPUser As String
        Public sSAPPwd As String
        Public sSAPDBName As String
        Public sDBUser As String
        Public sDBPwd As String

        Public sInboxDir As String
        Public sSuccessDir As String
        Public sFailDir As String
        Public sLogPath As String
        Public sDebug As String


        'Email Credentials

        Public sSMTPServer As String
        Public sSMTPPort As String
        Public sEmailFrom As String
        Public sSMTPUser As String
        Public sSMTPPassword As String
        Public sToEmailID As String
       
    End Structure


    ' Return Value Variable Control
    Public Const RTN_SUCCESS As Int16 = 1
    Public Const RTN_ERROR As Int16 = 0
    ' Debug Value Variable Control
    Public Const DEBUG_ON As Int16 = 1
    Public Const DEBUG_OFF As Int16 = 0

    ' Global variables group
    Public p_iDebugMode As Int16 = DEBUG_ON
    Public p_iErrDispMethod As Int16
    Public p_iDeleteDebugLog As Int16
    Public p_oCompDef As CompanyDefault
    Public p_dProcessing As DateTime
    Public p_oDtSuccess As DataTable
    Public p_oDtError As DataTable
    Public p_SyncDateTime As String
    Public p_oCompany As SAPbobsCOM.Company

    Public p_sSAPEntityName As String = String.Empty
    Public p_sSAPUName As String = String.Empty
    Public p_sSAPUPass As String = String.Empty
    Public p_iPWCrowCount As Integer = 0

    'Global DataTable

    Public p_oEntitesDetails As DataTable
    Public p_oSTOLDCODE As DataTable
    Public p_oDTPWCRowCount As DataTable
    Public P_sQueryString As String = String.Empty
    Public oDT_OUCODE As DataTable = Nothing


#End Region

    Sub Main()

        Dim sFuncName As String = String.Empty
        Dim sErrDesc As String = String.Empty

        Try
            sFuncName = "Main"

            Console.WriteLine("Starting Function", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            Console.WriteLine("Calling GetSystemIntializeInfo() ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling GetSystemIntializeInfo()", sFuncName)
            If GetSystemIntializeInfo(p_oCompDef, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            ''Console.WriteLine("Calling UDT_UDF_Creation() ", sFuncName)
            ''If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling UDT_UDF_Creation()", sFuncName)
            ''If UDT_UDF_Creation(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)


            Console.WriteLine("Calling GetEntitiesDetails() ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling GetEntitiesDetails() ", sFuncName)
            If GetEntitiesDetails(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            Console.WriteLine("Calling IdentifyTXTFile_JournalEntry() ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling IdentifyTXTFile_JournalEntry() ", sFuncName)
            If IdentifyExcelFile(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            Console.WriteLine("Completed With SUCCESS ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With SUCCESS ", sFuncName)

        Catch ex As Exception
            Console.WriteLine("Completed with ERROR ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
        End Try

    End Sub

End Module

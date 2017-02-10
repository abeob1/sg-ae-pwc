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
        Public sServer As String
        Public sLicenseServer As String
        Public sDBName As String
        Public sServerType As String
        Public sSAPUser As String
        Public sSAPPwd As String
        Public sSAPDBName As String
        Public sDBUser As String
        Public sDBPwd As String

        Public sDebug As String
        Public sFilepath As String

        'Email Credentials

        Public sSMTPServer As String
        Public sSMTPPort As String
        Public sEmailFrom As String
        Public sSMTPUser As String
        Public sSMTPPassword As String
        Public sToEmailID As String
        Public sSSL As String
    End Structure

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



End Module






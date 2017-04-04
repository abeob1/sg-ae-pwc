Option Strict Off
Option Explicit On
Friend Class HelloWorld
 
    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private oCompany As SAPbobsCOM.Company
    Public SboGuiApi As New SAPbouiCOM.SboGuiApi
    Public sConnectionString As String
    Dim oF_OSS_Report As OSS_Report

    Private Sub SetApplication()

        '*******************************************************************
        '// Use an SboGuiApi object to establish connection
        '// with the SAP Business One application and return an
        '// initialized appliction object
        '*******************************************************************

        Dim SboGuiApi As SAPbouiCOM.SboGuiApi
        Dim sConnectionString As String

        SboGuiApi = New SAPbouiCOM.SboGuiApi

        '// by following the steps specified above, the following
        '// statment should be suficient for either development or run mode

        sConnectionString = Environment.GetCommandLineArgs.GetValue(1)

        '// connect to a running SBO Application

        SboGuiApi.Connect(sConnectionString)

        '// get an initialized application object

        SBO_Application = SboGuiApi.GetApplication()

    End Sub

    Private Function SetConnectionContext() As Integer

        Dim sCookie As String
        Dim sConnectionContext As String
        Dim lRetCode As Integer

        '// First initialize the Company object

        oCompany = New SAPbobsCOM.Company

        '// Acquire the connection context cookie from the DI API.
        sCookie = oCompany.GetContextCookie

        '// Retrieve the connection context string from the UI API using the
        '// acquired cookie.
        sConnectionContext = SBO_Application.Company.GetConnectionContext(sCookie)

        '// before setting the SBO Login Context make sure the company is not
        '// connected

        If oCompany.Connected = True Then
            oCompany.Disconnect()
        End If

        '// Set the connection context information to the DI API.
        SetConnectionContext = oCompany.SetSboLoginContext(sConnectionContext)

    End Function

    Private Function ConnectToCompany() As Integer

        '// Establish the connection to the company database.
        ConnectToCompany = oCompany.Connect

    End Function
    'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'
    Private Sub Class_Initialize_Renamed()

        '//*************************************************************
        '// set SBO_Application with an initialized application object
        '//*************************************************************

        SetApplication()

        '//*************************************************************
        '// Set The Connection Context
        '//*************************************************************

        If Not SetConnectionContext() = 0 Then
            SBO_Application.MessageBox("Failed setting a connection to DI API")
            End ' Terminating the Add-On Application
        End If


        '//*************************************************************
        '// Connect To The Company Data Base
        '//*************************************************************

        If Not ConnectToCompany() = 0 Then
            SBO_Application.MessageBox("Failed connecting to the company's Data Base")
            End ' Terminating the Add-On Application
        End If

        '//*************************************************************
        '// send an "hello world" message
        '//*************************************************************
        SBO_Application.StatusBar.SetText("Addon Connected", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        '  SBO_Application.MessageBox("DI Connected To: " & oCompany.CompanyName & vbNewLine & "Hello World!")

    End Sub
    Public Sub New()
        MyBase.New()
        Try
            Class_Initialize_Renamed()
            LoadFromXML("MyMenus.xml", SBO_Application)
            oF_OSS_Report = New OSS_Report(oCompany, SBO_Application)
            SBO_Application.StatusBar.SetText("Addon Loaded", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#Region "Connection"
    'Public Sub New()
    '    MyBase.new()
    '    Try
    '        conn2()
    '        LoadFromXML("MyMenus.xml", SBO_Application)
    '        oF_OSS_Report = New OSS_Report(oCompany, SBO_Application)

    '        SBO_Application.StatusBar.SetText("Addon Loaded", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
    '    Catch ex As Exception
    '        SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '    End Try
    'End Sub
    Public Sub LoadFromXML(ByVal FileName As String, ByVal Sbo_application As SAPbouiCOM.Application)
        Dim oXmlDoc As New Xml.XmlDocument
        Dim sPath As String
        sPath = IO.Directory.GetParent(Application.StartupPath).ToString
        oXmlDoc.Load(sPath & "\PWC\" & FileName)
        Sbo_application.LoadBatchActions(oXmlDoc.InnerXml)
    End Sub
    'Public Sub conn1()
    '    Dim sconn As String
    '    Dim ret As Integer
    '    Dim scook As String
    '    Dim str As String
    '    Try
    '        sconn = Environment.GetCommandLineArgs.GetValue(1)
    '        SboGuiApi.Connect(sconn)
    '        SBO_Application = SboGuiApi.GetApplication
    '        SboGuiApi = Nothing
    '        scook = ocompany.GetContextCookie
    '        str = SBO_Application.Company.GetConnectionContext(scook)
    '        ret = ocompany.SetSboLoginContext(str)
    '        ocompany.Connect()
    '        ocompany.GetLastError(ret, str)
    '    Catch ex As Exception
    '        SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '    End Try
    'End Sub

#End Region
    Private Sub SBO_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles SBO_Application.AppEvent
        Try
            If EventType = SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged Then
                SBO_Application.StatusBar.SetText("Shuting Down addon", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Windows.Forms.Application.Exit()
            End If

            If EventType = SAPbouiCOM.BoAppEventTypes.aet_ShutDown Then
                SBO_Application.StatusBar.SetText("Shuting Down addon", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Windows.Forms.Application.Exit()
            End If

            If EventType = SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition Then
                SBO_Application.StatusBar.SetText("Shuting Down addon", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Windows.Forms.Application.Exit()

            End If

        Catch ex As Exception
            'Functions.WriteLog("Class:Connection" + " Function:SBO_Application_AppEvent" + " Error Message:" + ex.ToString)
        End Try
    End Sub
    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        Try
            If pVal.BeforeAction = True Then
                If pVal.MenuUID = "FMMySubMenu02" Then
                    LoadFromXML("OSS.srf", SBO_Application)
                    oForm = SBO_Application.Forms.Item("OSS")
                    oF_OSS_Report.Form_Bind(oForm)
                End If

            End If
        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
End Class
Imports System
Imports System.Diagnostics
Imports System.IO
Imports System.Security.Principal
Imports System.ServiceProcess
Imports System.Text
Imports System.Threading

Public Class ServiceController
    Public Enum ServiceStartupType
        Manual
        Automatic
        AutomaticDelayed
    End Enum

    Public DependOnServices As String()
    Public Description As String
    Public DisplayName As String
    Public ServerName As String
    Public ServiceName As String
    Public StartupType As ServiceStartupType

    Public Sub New(ByVal ServiceName As String)
        Me.ServiceName = ServiceName
        Me.ServerName = Environment.MachineName
        Me.StartupType = ServiceStartupType.Manual
    End Sub

    Public Sub New(ByVal ServiceName As String, ByVal ServerName As String)
        Me.ServiceName = ServiceName
        Me.ServerName = ServerName
        Me.StartupType = ServiceStartupType.Manual
    End Sub

    Private Function ExecuteBatchContain(ByVal TextContain As String) As String
        Try
            Dim TempFolder As String = Path.GetTempPath()
            Dim sw As New StreamWriter(Path.Combine(TempFolder, "ServiceController.bat"))
            sw.Write(TextContain)
            sw.Close()

            Dim oProcess As New Process()
            With oProcess.StartInfo
                .CreateNoWindow = True
                .WindowStyle = ProcessWindowStyle.Hidden
                .UseShellExecute = True
                .WorkingDirectory = TempFolder
                .FileName = "ServiceController.bat"
                .Verb = "runas"
                .Arguments = "/env /user:Administrator cmd"
            End With

            oProcess.Start()
            oProcess.WaitForExit()
            oProcess.Close()
            Return ""
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function


    Public Function GetServiceController() As System.ServiceProcess.ServiceController
        Dim oSCs As System.ServiceProcess.ServiceController() = System.ServiceProcess.ServiceController.GetServices(Me.ServerName)
        For Each oSC As System.ServiceProcess.ServiceController In oSCs
            If oSC.ServiceName = Me.ServiceName Then
                Return oSC
            End If
        Next
        Return Nothing
    End Function

    Private Function IsRunAsAdministrator() As Boolean
        Thread.GetDomain().SetPrincipalPolicy(PrincipalPolicy.WindowsPrincipal)
        Dim myPrincipal As WindowsPrincipal = DirectCast(Thread.CurrentPrincipal, WindowsPrincipal)
        Return myPrincipal.IsInRole(WindowsBuiltInRole.Administrator)
    End Function

    Public Shared Sub OpenManagementConsole()
        Process.Start("services.msc")
    End Sub

    Public Function Pause() As String
        Try
            If Me.GetServiceController Is Nothing Then
                Throw New Exception(String.Format("Service {0} is not existed!", Me.ServiceName))
            End If
            Dim sReturn As String = Me.ExecuteBatchContain(String.Format("sc \\{0} pause ""{1}""", Me.ServerName, Me.ServiceName))
            Return ""
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function Register(ByVal BinPath As String) As String
        Dim sbBatchContain As New StringBuilder()
        sbBatchContain.AppendFormat("sc \\{0} stop ""{1}""", Me.ServerName, Me.ServiceName)
        sbBatchContain.AppendLine()
        sbBatchContain.AppendFormat("sc \\{0} delete ""{1}""", Me.ServerName, Me.ServiceName)
        sbBatchContain.AppendLine()
        sbBatchContain.AppendFormat("sc \\{0} create {1} binPath= ""{2}""", Me.ServerName, Me.ServiceName, BinPath)
        If Not String.IsNullOrEmpty(Me.DisplayName) Then
            sbBatchContain.AppendFormat(" DisplayName= ""{0}""", Me.DisplayName)
        End If
        Select Case Me.StartupType
            Case ServiceStartupType.Automatic
                sbBatchContain.Append(" start= auto")
                Exit Select

            Case ServiceStartupType.AutomaticDelayed
                sbBatchContain.Append(" start= delayed-auto")
                Exit Select

            Case ServiceStartupType.Manual
                Exit Select
            Case Else

                sbBatchContain.Append(" start= demand")
                Exit Select
        End Select
        If (Not Me.DependOnServices Is Nothing) AndAlso (Me.DependOnServices.Length > 0) Then
            Dim DependServiceNames As String = ""
            For Each DependServiceName As String In Me.DependOnServices
                If DependServiceNames = "" Then
                    DependServiceNames = DependServiceNames + DependServiceName
                Else
                    DependServiceNames = DependServiceNames + "/" + DependServiceName
                End If
            Next
            sbBatchContain.AppendFormat(" depend= ""{0}""", DependServiceNames)
        End If
        If Not String.IsNullOrEmpty(Me.Description) Then
            sbBatchContain.AppendLine()
            sbBatchContain.AppendFormat("sc \\{0} description {1} ""{2}""", Me.ServerName, Me.ServiceName, Me.Description)
        End If
        'Functions.WriteLog(sbBatchContain.ToString())
        Dim sReturn As String = Me.ExecuteBatchContain(sbBatchContain.ToString())
        If sReturn <> "" Then
            Return sReturn
        End If
        Dim oSC As System.ServiceProcess.ServiceController = Me.GetServiceController()
        If Me.GetServiceController Is Nothing Then
            Return String.Format("Cannot create service {0}. Please create it manually!", Me.ServiceName)
        End If
        Return ""
    End Function

    Public Function [Resume]() As String
        Try
            If Me.GetServiceController Is Nothing Then
                Throw New Exception(String.Format("Service {0} is not existed!", Me.ServiceName))
            End If
            Dim sReturn As String = Me.ExecuteBatchContain(String.Format("sc \\{0} continue ""{1}""", Me.ServerName, Me.ServiceName))
            Return ""
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function Start() As String
        Try
            If Me.GetServiceController Is Nothing Then
                Throw New Exception(String.Format("Service {0} is not existed!", Me.ServiceName))
            End If
            Dim sReturn As String = Me.ExecuteBatchContain(String.Format("sc \\{0} start ""{1}""", Me.ServerName, Me.ServiceName))
            Return ""
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function [Stop]() As String
        Try
            If Me.GetServiceController Is Nothing Then
                Throw New Exception(String.Format("Service {0} is not existed!", Me.ServiceName))
            End If
            Dim sReturn As String = Me.ExecuteBatchContain(String.Format("sc \\{0} stop ""{1}""", Me.ServerName, Me.ServiceName))
            Return ""
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function Unregister() As String
        If Not Me.GetServiceController Is Nothing Then
            Dim sbBatchContain As New StringBuilder()
            sbBatchContain.AppendFormat("sc \\{0} stop ""{1}""", Me.ServerName, Me.ServiceName)
            sbBatchContain.AppendLine()
            sbBatchContain.AppendFormat("sc \\{0} delete ""{1}""", Me.ServerName, Me.ServiceName)
            Dim sReturn As String = Me.ExecuteBatchContain(sbBatchContain.ToString())
            If sReturn <> "" Then
                Return sReturn
            End If
        End If
        If Me.GetServiceController Is Nothing Then
            Return ""
        End If
        Return String.Format("Cannot delete service {0}. Please delete it manually!", Me.ServiceName)
    End Function

    Public ReadOnly Property Status() As String
        Get
            Dim oSC As System.ServiceProcess.ServiceController = Me.GetServiceController()
            If oSC Is Nothing Then
                Return ""
            End If
            Return oSC.Status.ToString()
        End Get
    End Property
End Class

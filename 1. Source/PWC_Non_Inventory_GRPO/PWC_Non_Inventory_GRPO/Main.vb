Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Windows.Forms
Imports System.ServiceProcess
Module SubMain
    Public Sub Main()

        Dim Args As String() = Environment.GetCommandLineArgs()

        If Args.Length = 1 Then
            Application.EnableVisualStyles()
            Application.SetCompatibleTextRenderingDefault(False)
            Application.Run(New frmMornitor())
        ElseIf Args(1) = "-service" Then
            Dim ServicesToRun As ServiceBase()
            ServicesToRun = New ServiceBase() {New Service1()}
            ServiceBase.Run(ServicesToRun)
        End If
    End Sub
End Module


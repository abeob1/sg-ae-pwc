

Public Class Service1
#Region "General"

    Protected Overrides Sub OnStart(ByVal args() As String)
        Try
            Timer1.Enabled = True
            Timer1.Interval = System.Configuration.ConfigurationSettings.AppSettings.Get("Timer") * 60000
            Timer1.Start()

            Timer3.Enabled = True
            Timer3.Start()
        Catch ex As Exception
            Functions.WriteLog(ex.ToString)
        End Try
    End Sub
    Protected Overrides Sub OnStop()
        Try
            Timer1.Enabled = False
            Timer1.Stop()

            Timer3.Enabled = False
            Timer3.Stop()
        Catch ex As Exception
            Functions.WriteLog(ex.ToString)
        End Try
    End Sub
    Private Sub Timer1_Elapsed(ByVal sender As System.Object, ByVal e As System.Timers.ElapsedEventArgs) Handles Timer1.Elapsed
        Try
            'Dim xm As New oXML
            'xm.SetDB()
            'Dim oerrm As New oEmailError
            'oerrm.SendErrorEmail()

            Timer1.Enabled = False
            Dim a As New Functions
            a.AutoRun()
            Timer1.Enabled = True
        Catch ex As Exception
            Functions.WriteLog(ex.ToString)
            Timer1.Enabled = True
        End Try
    End Sub
#End Region
    
    Private Sub Timer3_Elapsed(sender As System.Object, e As System.Timers.ElapsedEventArgs) Handles Timer3.Elapsed
        Try
            '60 minutes
            'Functions.WriteLog("Start Timer 2")
            Dim xm As New oXML
            xm.SetDB()
            '------------SEND ERROR EMAIL-------------
            If Integer.Parse(DateTime.Now.ToString("HH")) = 8 Then
                'Functions.WriteLog("Compare 8")
                Dim oerrm As New oEmailError
                oerrm.SendErrorEmail()
            End If
            'Functions.WriteLog("End Timer 2")
        Catch ex As Exception
            Functions.WriteLog(ex.ToString)
        End Try
    End Sub
End Class

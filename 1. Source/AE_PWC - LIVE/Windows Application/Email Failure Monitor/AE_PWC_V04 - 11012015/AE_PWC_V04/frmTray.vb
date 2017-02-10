Public Class frmTray

    Public sErrDesc As String = String.Empty
    Public sFuncName As String = String.Empty



    Private Sub Form1_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        e.Cancel = True
        Me.WindowState = FormWindowState.Minimized
        Me.ShowInTaskbar = False
        CmenuStrip.Enabled = True

    End Sub


    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.Visible = False
        CmenuStrip.Enabled = True
        sFuncName = "Tray Load()"
        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling GetSystemIntializeInfo()", sFuncName)
        If GetSystemIntializeInfo(p_oCompDef, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

    End Sub


    Private Sub EmailErrorMonitoringToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles EmailErrorMonitoringToolStripMenuItem.Click
        ''Me.WindowState = FormWindowState.Normal
        ''Me.ShowInTaskbar = True
        '' CmenuStrip.Enabled = False
        EmailMonitor_ShowDialog()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        End
    End Sub

    Private Sub Form1_SizeChanged(sender As Object, e As System.EventArgs) Handles Me.SizeChanged
        If Me.WindowState = FormWindowState.Minimized Then
            ShowInTaskbar = False
            CmenuStrip.Enabled = True
        End If

    End Sub

    Private Sub Nicon_MouseDoubleClick(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles Nicon.MouseDoubleClick

    End Sub
End Class

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTray
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmTray))
        Me.CmenuStrip = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.EmailErrorMonitoringToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Nicon = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.CmenuStrip.SuspendLayout()
        Me.SuspendLayout()
        '
        'CmenuStrip
        '
        Me.CmenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.EmailErrorMonitoringToolStripMenuItem, Me.ToolStripSeparator1, Me.ExitToolStripMenuItem, Me.ToolStripSeparator2})
        Me.CmenuStrip.Name = "CmenuStrip"
        Me.CmenuStrip.Size = New System.Drawing.Size(188, 82)
        '
        'EmailErrorMonitoringToolStripMenuItem
        '
        Me.EmailErrorMonitoringToolStripMenuItem.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.EmailErrorMonitoringToolStripMenuItem.Image = Global.AE_PWC_V04.My.Resources.Resources.Monitor
        Me.EmailErrorMonitoringToolStripMenuItem.Name = "EmailErrorMonitoringToolStripMenuItem"
        Me.EmailErrorMonitoringToolStripMenuItem.Size = New System.Drawing.Size(187, 22)
        Me.EmailErrorMonitoringToolStripMenuItem.Text = "Email Failure Monitor"
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Image = Global.AE_PWC_V04.My.Resources.Resources._Exit
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(187, 22)
        Me.ExitToolStripMenuItem.Text = "Exit"
        '
        'Nicon
        '
        Me.Nicon.ContextMenuStrip = Me.CmenuStrip
        Me.Nicon.Icon = CType(resources.GetObject("Nicon.Icon"), System.Drawing.Icon)
        Me.Nicon.Text = "Email Failure Monitor"
        Me.Nicon.Visible = True
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(184, 6)
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(184, 6)
        '
        'frmTray
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 261)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmTray"
        Me.ShowInTaskbar = False
        Me.Text = "Email Error Notification"
        Me.WindowState = System.Windows.Forms.FormWindowState.Minimized
        Me.CmenuStrip.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents CmenuStrip As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents EmailErrorMonitoringToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Nicon As System.Windows.Forms.NotifyIcon
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripSeparator2 As System.Windows.Forms.ToolStripSeparator

End Class

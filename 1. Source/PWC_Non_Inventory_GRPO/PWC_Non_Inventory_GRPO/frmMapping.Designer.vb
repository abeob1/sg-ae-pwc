<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMapping
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMapping))
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.btnStop = New System.Windows.Forms.Button()
        Me.btnStart = New System.Windows.Forms.Button()
        Me.grMonitor = New System.Windows.Forms.DataGridView()
        Me.Panel2.SuspendLayout()
        CType(Me.grMonitor, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.btnStop)
        Me.Panel2.Controls.Add(Me.btnStart)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel2.Location = New System.Drawing.Point(0, 468)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(658, 39)
        Me.Panel2.TabIndex = 2
        '
        'btnStop
        '
        Me.btnStop.Location = New System.Drawing.Point(84, 5)
        Me.btnStop.Name = "btnStop"
        Me.btnStop.Size = New System.Drawing.Size(75, 30)
        Me.btnStop.TabIndex = 1
        Me.btnStop.Text = "Close"
        Me.btnStop.UseVisualStyleBackColor = True
        '
        'btnStart
        '
        Me.btnStart.Location = New System.Drawing.Point(3, 5)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(75, 30)
        Me.btnStart.TabIndex = 0
        Me.btnStart.Text = "Save"
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'grMonitor
        '
        Me.grMonitor.AllowUserToAddRows = False
        Me.grMonitor.AllowUserToDeleteRows = False
        Me.grMonitor.AllowUserToOrderColumns = True
        Me.grMonitor.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grMonitor.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grMonitor.Location = New System.Drawing.Point(0, 0)
        Me.grMonitor.Name = "grMonitor"
        Me.grMonitor.Size = New System.Drawing.Size(658, 468)
        Me.grMonitor.TabIndex = 5
        '
        'frmMapping
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(658, 507)
        Me.Controls.Add(Me.grMonitor)
        Me.Controls.Add(Me.Panel2)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmMapping"
        Me.Text = "Mapping Table (Ver. 2013.02.22)"
        Me.Panel2.ResumeLayout(False)
        CType(Me.grMonitor, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents btnStop As System.Windows.Forms.Button
    Friend WithEvents btnStart As System.Windows.Forms.Button
    Friend WithEvents grMonitor As System.Windows.Forms.DataGridView
End Class

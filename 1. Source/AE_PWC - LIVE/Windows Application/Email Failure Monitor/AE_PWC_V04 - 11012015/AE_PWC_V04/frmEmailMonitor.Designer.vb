<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmEmailMonitor
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmEmailMonitor))
        Me.btnRefresh = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.labmsg = New System.Windows.Forms.Label()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.Panel5 = New System.Windows.Forms.Panel()
        Me.dvrEmailmointor = New System.Windows.Forms.DataGridView()
        Me.Choose = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.DocType = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Draftkey = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.entity = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.emailid = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.errmsg = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Refno = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Panel3.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.Panel5.SuspendLayout()
        CType(Me.dvrEmailmointor, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnRefresh
        '
        Me.btnRefresh.Location = New System.Drawing.Point(168, 9)
        Me.btnRefresh.Name = "btnRefresh"
        Me.btnRefresh.Size = New System.Drawing.Size(136, 23)
        Me.btnRefresh.TabIndex = 2
        Me.btnRefresh.Text = "Refresh"
        Me.btnRefresh.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(87, 9)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 1
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnUpdate
        '
        Me.btnUpdate.Location = New System.Drawing.Point(6, 9)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(75, 23)
        Me.btnUpdate.TabIndex = 0
        Me.btnUpdate.Text = "Update"
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Left
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(29, 463)
        Me.Panel1.TabIndex = 3
        '
        'Panel2
        '
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Location = New System.Drawing.Point(29, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(1221, 23)
        Me.Panel2.TabIndex = 4
        '
        'Panel3
        '
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Controls.Add(Me.Label1)
        Me.Panel3.Controls.Add(Me.labmsg)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel3.Location = New System.Drawing.Point(29, 436)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(1221, 27)
        Me.Panel3.TabIndex = 5
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(5, 4)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(33, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Msg :"
        '
        'labmsg
        '
        Me.labmsg.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labmsg.Location = New System.Drawing.Point(42, 3)
        Me.labmsg.Name = "labmsg"
        Me.labmsg.Size = New System.Drawing.Size(537, 20)
        Me.labmsg.TabIndex = 0
        '
        'Panel4
        '
        Me.Panel4.Controls.Add(Me.btnUpdate)
        Me.Panel4.Controls.Add(Me.btnCancel)
        Me.Panel4.Controls.Add(Me.btnRefresh)
        Me.Panel4.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel4.Location = New System.Drawing.Point(29, 397)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(1221, 39)
        Me.Panel4.TabIndex = 6
        '
        'Panel5
        '
        Me.Panel5.Controls.Add(Me.dvrEmailmointor)
        Me.Panel5.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel5.Location = New System.Drawing.Point(29, 23)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(1221, 374)
        Me.Panel5.TabIndex = 7
        '
        'dvrEmailmointor
        '
        Me.dvrEmailmointor.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dvrEmailmointor.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Choose, Me.DocType, Me.Draftkey, Me.entity, Me.emailid, Me.errmsg, Me.Refno})
        Me.dvrEmailmointor.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dvrEmailmointor.GridColor = System.Drawing.SystemColors.ControlLight
        Me.dvrEmailmointor.Location = New System.Drawing.Point(0, 0)
        Me.dvrEmailmointor.Name = "dvrEmailmointor"
        Me.dvrEmailmointor.RowHeadersVisible = False
        Me.dvrEmailmointor.Size = New System.Drawing.Size(1221, 374)
        Me.dvrEmailmointor.TabIndex = 3
        '
        'Choose
        '
        Me.Choose.HeaderText = "Choose"
        Me.Choose.Name = "Choose"
        Me.Choose.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Choose.Width = 60
        '
        'DocType
        '
        Me.DocType.HeaderText = "Document Type"
        Me.DocType.Name = "DocType"
        Me.DocType.ReadOnly = True
        Me.DocType.Width = 150
        '
        'Draftkey
        '
        Me.Draftkey.HeaderText = "Draft Key"
        Me.Draftkey.Name = "Draftkey"
        Me.Draftkey.ReadOnly = True
        '
        'entity
        '
        Me.entity.HeaderText = "Entity"
        Me.entity.Name = "entity"
        Me.entity.ReadOnly = True
        '
        'emailid
        '
        Me.emailid.HeaderText = "Email Address"
        Me.emailid.Name = "emailid"
        Me.emailid.Width = 400
        '
        'errmsg
        '
        Me.errmsg.HeaderText = "Error Msg"
        Me.errmsg.Name = "errmsg"
        Me.errmsg.ReadOnly = True
        Me.errmsg.Width = 300
        '
        'Refno
        '
        Me.Refno.HeaderText = "Ref No."
        Me.Refno.Name = "Refno"
        Me.Refno.ReadOnly = True
        '
        'Timer1
        '
        Me.Timer1.Interval = 3000
        '
        'frmEmailMonitor
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1250, 463)
        Me.Controls.Add(Me.Panel5)
        Me.Controls.Add(Me.Panel4)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmEmailMonitor"
        Me.ShowInTaskbar = False
        Me.Text = "Email Rectification "
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.Panel4.ResumeLayout(False)
        Me.Panel5.ResumeLayout(False)
        CType(Me.dvrEmailmointor, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnRefresh As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents labmsg As System.Windows.Forms.Label
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents Panel5 As System.Windows.Forms.Panel
    Friend WithEvents dvrEmailmointor As System.Windows.Forms.DataGridView
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents Choose As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents DocType As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Draftkey As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents entity As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents emailid As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents errmsg As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Refno As System.Windows.Forms.DataGridViewTextBoxColumn

End Class

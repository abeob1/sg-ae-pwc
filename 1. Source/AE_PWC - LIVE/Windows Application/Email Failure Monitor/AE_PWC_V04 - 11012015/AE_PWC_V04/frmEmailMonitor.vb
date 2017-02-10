Imports System.IO
Imports Microsoft.VisualBasic

Public Class frmEmailMonitor

    Dim sErrDesc As String = String.Empty


    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs)
        Me.Close()
    End Sub


    Private Sub frmEmailMonitor_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "Load_Emailstatusfails()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Load_Emailstatusfails()", sFuncName)
            If Load_Emailstatusfails(sErrDesc, Me.dvrEmailmointor) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        Catch ex As Exception

        End Try

    End Sub

    Private Sub btnCancel_Click(sender As System.Object, e As System.EventArgs) Handles btnCancel.Click
        Me.Dispose()

    End Sub

    Private Sub btnUpdate_Click(sender As System.Object, e As System.EventArgs) Handles btnUpdate.Click

        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "Update_Emailmonitor"
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.btnUpdate.Enabled = False
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling function Update_Emailstatusfails()", sFuncName)
            If Update_Emailstatusfails(sErrDesc, Me.dvrEmailmointor) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            Call btnRefresh_Click(Me, New System.EventArgs)
            Me.btnUpdate.Enabled = True
            System.Windows.Forms.Cursor.Current = Cursors.Default
            Me.labmsg.Text = "Operation Completed Successfully"
            Timer1.Enabled = True
        Catch ex As Exception

        End Try

    End Sub

    Private Sub btnRefresh_Click(sender As System.Object, e As System.EventArgs) Handles btnRefresh.Click
        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "Refresh_Emailstatusfails()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
            If Load_Emailstatusfails(sErrDesc, Me.dvrEmailmointor) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            Me.labmsg.Text = "Operation Completed Successfully "
            Timer1.Enabled = True
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        Me.labmsg.Text = String.Empty
        Timer1.Enabled = False
    End Sub

    Private Sub frmEmailMonitor_Resize(sender As Object, e As System.EventArgs) Handles Me.Resize
        If Me.WindowState = FormWindowState.Minimized Then
            Me.Dispose()
        End If
    End Sub

 
    Private Sub dvrEmailmointor_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dvrEmailmointor.CellContentClick
        If dvrEmailmointor.Columns.Item(e.ColumnIndex).Name = "Choose" And e.RowIndex = -1 Then
            Dim bflag As Boolean = False
            If Convert.ToBoolean(dvrEmailmointor.Rows(0).Cells(0).Value) = True Then
                bflag = False
            Else
                bflag = True
            End If

            Me.dvrEmailmointor.Rows(0).Cells(2).Selected = True
            For imjs As Integer = 0 To Me.dvrEmailmointor.Rows.Count - 2 'oDsBPList.Tables(0).Rows.Count - 1
                Me.dvrEmailmointor.Rows.Item(imjs).Cells.Item("Choose").Value = bflag  'oDsBPList.Tables(0).Rows(imjs)("Check").ToString
            Next
            Me.dvrEmailmointor.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize

            btnUpdate.Focus()

        End If
    End Sub
End Class

Imports System.IO

Public Class frmMornitor
#Region "Service"
    Private Sub RefreshStatus()
        Dim a As New ServiceController("SAPIntegration")
        If a.Status = "" Then
            btnReg.Enabled = True
            btnUnReg.Enabled = False
            btnStart.Enabled = False
            btnStop.Enabled = False
        ElseIf a.Status = "Stopped" Then
            btnStart.Enabled = True
            btnStop.Enabled = False
            btnReg.Enabled = False
            btnUnReg.Enabled = True
        ElseIf a.Status = "Running" Then
            btnStart.Enabled = False
            btnStop.Enabled = True
            btnReg.Enabled = False
            btnUnReg.Enabled = True
        End If
    End Sub
    Private Sub btnStart_Click(sender As System.Object, e As System.EventArgs) Handles btnStart.Click
        Dim a As New ServiceController("SAPIntegration")
        Dim str As String
        str = a.Start()
        RefreshStatus()
    End Sub
    Private Sub btnStop_Click(sender As System.Object, e As System.EventArgs) Handles btnStop.Click
        Dim a As New ServiceController("SAPIntegration")
        Dim str As String = a.Stop()
        RefreshStatus()
    End Sub
    Private Sub btnReg_Click(sender As System.Object, e As System.EventArgs) Handles btnReg.Click
        Dim a As New ServiceController("SAPIntegration")
        a.Description = "SAPIntegration"
        a.DisplayName = "SAPIntegration"
        a.ServiceName = "SAPIntegration"
        a.StartupType = ServiceController.ServiceStartupType.Automatic

        Dim sReturn As String
        sReturn = a.Register(Application.ExecutablePath + " -service")
        If sReturn = "" Then
            MessageBox.Show("Register Sucessfull! - Application will be closed.")
            Application.Exit()
        Else
            MessageBox.Show("Error: " + sReturn)
        End If
        RefreshStatus()
    End Sub
    Private Sub btnUnReg_Click(sender As System.Object, e As System.EventArgs) Handles btnUnReg.Click
        Dim a As New ServiceController("SAPIntegration")
        Dim sReturn As String
        sReturn = a.Unregister()
        If sReturn = "" Then
            MessageBox.Show("UnRegister Sucessfull!")
        Else
            MessageBox.Show("Error: " + sReturn)
        End If
        RefreshStatus()
    End Sub
#End Region
#Region "Events"
    Private Sub frmMornitor_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Try
            ' Exit Sub
            Dim fn As New oXML
            fn.SetDB()
            LoadComboBox_Status()
            LoadComboBox_DB()
            RefreshMonitor()
            RefreshStatus()

        Catch ex As Exception
            MessageBox.Show("Load" + ex.ToString)
            Me.Cursor = Cursors.Default
            Functions.WriteLog("frmMornitor_Load:" & ex.Message)
        End Try
    End Sub
    Public Sub LoadComboBox_Status()
        Try
            ComboBox2.SelectedIndex = 0
            Dim Str As String = ""
            Dim cn As New Connection
            Str = "select 'All' St union All select 'Pending' St union All select 'Successfull' St union All select 'Failed' St"
            Dim dt As DataTable = cn.Integration_RunQuery(Str)
            ComboBox1.DisplayMember = "St"
            ComboBox1.ValueMember = "St"
            ComboBox1.DataSource = dt
            ComboBox1.SelectedIndex = 0

        Catch ex As Exception
            Functions.WriteLog("LoadComboBox_Status:" & ex.Message)
        End Try
    End Sub
    Public Sub LoadComboBox_DB()
        Try
        
            Dim Str As String = ""
            Str = "SELECT T0.Name FROM [dbo].[@AB_COMPANYDATA]  T0 order by T0.Name"
            Dim cn As New Connection
            Dim dt As DataTable = cn.Integration_RunQuery(Str)
            ComboBox3.DisplayMember = "Name"
            ComboBox3.ValueMember = "Name"
            ComboBox3.DataSource = dt
            ComboBox3.SelectedIndex = 0


        Catch ex As Exception
            Functions.WriteLog("LoadComboBox_DB:" & ex.Message)
        End Try
    End Sub
    Private Sub btnRefresh_Click(sender As System.Object, e As System.EventArgs) Handles btnRefresh.Click
        RefreshMonitor()
    End Sub
    Private Sub btnClose_Click(sender As System.Object, e As System.EventArgs) Handles btnClose.Click
        Me.Close()
        Application.Exit()
    End Sub
    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        Timer1.Enabled = False
        Dim orunning As New oRunningMonitor
        lblRunning.Text = orunning.GetLastRunning()
        If ckAutoRef.Checked Then
            RefreshMonitor()
        End If
        Timer1.Enabled = True
    End Sub
    Private Sub btnLog_Click(sender As System.Object, e As System.EventArgs) Handles btnLog.Click
        Dim LogFileName As String = Application.StartupPath + "\logfile.txt"
        If File.Exists(LogFileName) Then
            System.Diagnostics.Process.Start(LogFileName)
        End If
    End Sub
    Private Sub btnRetryAll_Click(sender As System.Object, e As System.EventArgs) Handles btnRetryAll.Click
        Try

            Dim str As String = ""
        Dim opt1 As String = cbFilter.Text
        Dim strEnd As String = " set ReceiveDate=null, ErrMsg=null where ReceiveDate is not null and isnull(errMsg,'')<>''"
        Select Case opt1
            Case "GRPO"
                str = "Update [AB_GRPO_NON_INV]" + strEnd
            Case "Send Email"
                str = "Update SendEmailPOHeader" + strEnd
            End Select
            Dim DBName As String = ""
            If Me.ComboBox3.Items.Count <> 0 Then
                DBName = Me.ComboBox3.SelectedValue 'ComboBox3.SelectedItem.ToString()
            Else
                Exit Try
            End If

        Dim cn As New Connection
            Dim dt As DataTable = cn.Integration_RunQuery_BR(str, DBName)
            RefreshMonitor()
        Catch ex As Exception
            Functions.WriteLog("LoadComboBox_DB:" & ex.Message)
        End Try
    End Sub
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim frm As New frmMapping
        frm.ShowDialog()

    End Sub
    Private Sub btnRetry_Click(sender As System.Object, e As System.EventArgs) Handles btnRetry.Click
        Dim HeaderID As String = ""
        Dim HeaderIDFieldName As String = "ID"
        Try

            If cbFilter.Text = "Goods Receipt" Or cbFilter.Text = "Goods Issue" Then
                HeaderIDFieldName = "DocEntry"
            End If
            HeaderID = grMonitor.SelectedRows.Item(0).Cells(HeaderIDFieldName).Value

        Catch ex As Exception

        End Try

        Dim str As String = ""
        Dim opt1 As String = cbFilter.Text
        Dim strEnd As String = " set ReceiveDate=null, ErrMsg=null where ReceiveDate is not null and isnull(errMsg,'')<>'' and " + HeaderIDFieldName + "=" + CStr(grMonitor.SelectedRows.Item(0).Cells(HeaderIDFieldName).Value)

        Select Case opt1
            Case "Item"
                str = "Update ItemMasterData" + strEnd
            Case "Business Partner"
                str = "Update BusinessParterMaster" + strEnd
            Case "Purchase Order"
                str = "Update POHeader" + strEnd
            Case "GRPO"
                str = "Update GRPOHeader" + strEnd
            Case "Goods Return"
                str = "Update GoodsReturnHeader" + strEnd
            Case "Inventory Transfer"
                str = "Update TransferHeader" + strEnd
            Case "Invoice"
                str = "Update InvoiceHeader" + strEnd
            Case "Goods Receipt"
                str = "Update GoodsReceiptHeader" + strEnd
            Case "Goods Issue"
                str = "Update GoodsIssueHeader" + strEnd
            Case "Stock Take"
                str = "Update StockTake" + strEnd
            Case "Send Email"
                str = "Update SendEmailPOHeader" + strEnd

        End Select
        Dim cn As New Connection
        Dim dt As DataTable = cn.Integration_RunQuery(str)
        RefreshMonitor()
    End Sub
    Private Sub btnUpload_Click(sender As System.Object, e As System.EventArgs) Handles btnUpload.Click
        Try

            btnUpload.Enabled = False
        Dim frm As New oJournalEntry
        Dim MyArr As Array
        Dim Str As String
            Str = System.Configuration.ConfigurationSettings.AppSettings.Get("SAPConnectionString")
            Dim DBName As String = ""
            If Me.ComboBox3.Items.Count <> 0 Then
                DBName = Me.ComboBox3.SelectedValue 'ComboBox3.SelectedItem.ToString()
                MyArr = Str.Split(";")
                Dim constr As String = "Data Source= " + MyArr(3).ToString() + ";Initial Catalog=" + DBName + " ;User ID=" + MyArr(4).ToString() + "; Password=" + MyArr(5).ToString() + ";"
                frm.Insert_JE(constr, DBName)
                frm.CreateJE_LastMonth(constr, DBName)
                frm.CreateJE_FirstMonth(constr, DBName)
            End If
            btnUpload.Enabled = True
        Catch ex As Exception
            btnUpload.Enabled = True
            Functions.WriteLog("Manula Run btnUpload_Click Error: " & ex.Message)
        End Try
    End Sub
    Private Sub grMonitor_SelectionChanged(sender As System.Object, e As System.EventArgs)
        Dim HeaderID As String = ""
        Try
            Dim HeaderIDFieldName As String = "ID"
            If cbFilter.Text = "Goods Receipt" Or cbFilter.Text = "Goods Issue" Then
                HeaderIDFieldName = "DocEntry"
            End If
            HeaderID = grMonitor.SelectedRows.Item(0).Cells(HeaderIDFieldName).Value

        Catch ex As Exception

        End Try

        If HeaderID <> "" Then
            Dim strQuery As String = ""
            Dim cn As New Connection
            strQuery = BuildMonitorSubQuery(cbFilter.Text, HeaderID)
            If strQuery <> "" Then
                Dim dt As DataTable = cn.Integration_RunQuery(strQuery)
                grDetail.DataSource = dt
            Else
                grDetail.DataSource = Nothing
            End If
        End If
    End Sub
    Private Sub grMonitor_MouseDoubleClick(sender As System.Object, e As System.Windows.Forms.MouseEventArgs)
        Try
            Exit Sub
            If grMonitor.RowCount = 0 Then
                Return
            End If
            If cbFilter.Text = "Invoice" Then
                Dim frm As New frmPayment
                Dim cn As New Connection
                Dim strQuery As String = "select * from PaymentMean where HeaderID=" + CStr(grMonitor.SelectedRows.Item(0).Cells("ID").Value)
                Dim dt As DataTable = cn.Integration_RunQuery(strQuery)
                frm.grMonitor.DataSource = dt
                frm.ShowDialog()
            End If

            If cbFilter.Text = "Wincor Sales" Then
                Dim frm As New frmPayment
                Dim cn As New Connection
                Dim strQuery As String = "select * from WC_PaymentMean where HeaderID=" + CStr(grMonitor.SelectedRows.Item(0).Cells("ID").Value)
                Dim dt As DataTable = cn.Integration_RunQuery(strQuery)
                frm.grMonitor.DataSource = dt
                frm.ShowDialog()
            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub grDetail_MouseDoubleClick(sender As System.Object, e As System.Windows.Forms.MouseEventArgs)
        Try


            If grDetail.RowCount = 0 Then
                Return
            End If
            Dim frm As New frmSerial
            Dim cn As New Connection
            Dim transtype As String = ""
            Select Case cbFilter.Text
                Case "GRPO"
                    transtype = "20"
                Case "Goods Return"
                    transtype = "21"
                Case "Inventory Transfer"
                    transtype = "67"
                Case "Invoice"
                    transtype = "13"
                Case "Goods Receipt"
                    transtype = "59"
                Case "Goods Issue"
                    transtype = "60"
                Case "Stock Take"
                    transtype = "10000071"

            End Select
            Dim strQuery As String = "Select * from SerialNumber where HeaderID=" + CStr(grDetail.SelectedRows.Item(0).Cells("HeaderID").Value) + _
                                        " and TransType=" + transtype
            Dim dt As DataTable = cn.Integration_RunQuery(strQuery)
            frm.grMonitor.DataSource = dt
            frm.ShowDialog()
        Catch ex As Exception

        End Try
    End Sub
#End Region
#Region "Functions"
    Private Sub RefreshMonitor()
        Try

        
        Dim fn As New oXML
        fn.SetDB()

        Dim cn As New Connection
            Dim strQuery As String = ""
            If Me.ComboBox1.Items.Count <> 0 Then
                Dim Status As String = Me.ComboBox1.SelectedValue.ToString
                strQuery = BuildMonitorQuery(cbFilter.Text, Status)
            End If

            If Me.ComboBox3.Items.Count <> 0 Then
                Dim DBName As String = Me.ComboBox3.SelectedValue 'ComboBox3.SelectedItem.ToString()
                If DBName <> "" Then
                    Dim dt As DataTable = cn.Integration_RunQuery_BR(strQuery, DBName)
                    grMonitor.DataSource = dt
                End If
            End If
           


        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Function BuildMonitorQuery(opt1 As String, opt2 As String)
        '-----------option 2-----------
        'All
        'Pending
        'Successfull
        'Failed
        Dim strEnd As String = ""
        Select Case opt2
            Case "All"
                strEnd = ""
            Case "Pending"
                strEnd = " where [ReceiveDate_LastMonth] is null or [ReceiveDate_FitstMonth] is null"
            Case "Successfull"
                strEnd = " where ([ReceiveDate_LastMonth] is not null and isnull([ErrorMsg],'')='') or ([ReceiveDate_LastMonth] is not null and isnull([ErrorMsg1],'')='') "
            Case "Failed"
                strEnd = " where isnull([ErrorMsg],'')<>'' or isnull([ErrorMsg1],'')<>''"
            Case Else
                strEnd = ""
        End Select

        If cbSendDate.Checked Then
            If strEnd <> "" Then
                strEnd = strEnd + " AND "
            Else : strEnd = " WHERE "
            End If

            strEnd = strEnd + "  datediff(dd,SendDate,'" + cbSendDate.Value.ToString("MM/dd/yyyy") + "')=0"
        End If

        '-----------option 1-----------
        'GRPO
        Dim str As String = ""
        Select Case opt1
            Case "GRPO"
                str = "select Top 1000 * from [AB_GRPO_NON_INV]" + strEnd
            Case Else
                str = "select Top 1000 * from [AB_GRPO_NON_INV]" + strEnd
        End Select

        Return str
    End Function
    Private Function BuildMonitorSubQuery(opt1 As String, HeaderID As String)
        Return ""
        Exit Function
        Dim str As String = ""
        Select Case opt1
            Case "GRPO"
                str = "select * from GRPOLine where HeaderID=" + HeaderID
          
            Case Else
                str = ""
        End Select

        Return str
    End Function
#End Region

    
    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click

    End Sub

    Private Sub cbFilter_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cbFilter.SelectedIndexChanged
        RefreshMonitor()
    End Sub

   

    Private Sub cbResult_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cbResult.SelectedIndexChanged
        RefreshMonitor()
    End Sub

    Private Sub grMonitor_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs)

    End Sub

 

    Private Sub ComboBox1_SelectedIndexChanged_1(sender As System.Object, e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        RefreshMonitor()
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        RefreshMonitor()
    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        RefreshMonitor()
    End Sub
End Class
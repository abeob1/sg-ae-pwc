Public Class Service1

    Public oEmailTrigger As New System.Timers.Timer



    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Add code here to start your service. This method should set things
        ' in motion so your service can do its work.
        Try
            p_iDebugMode = DEBUG_ON

            sFuncName = "Onstart()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling GetSystemIntializeInfo() ", sFuncName)
            If GetSystemIntializeInfo(p_oCompDef, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Email Trigger Service Starts  " & Format(Now.Date, "dd-MMM-yyyy"), sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("----------------------------------------------------------------------", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Timer Activated  ", sFuncName)
            oEmailTrigger.Interval = 300000
            oEmailTrigger.Start()
            AddHandler oEmailTrigger.Elapsed, AddressOf EmailNotifiation

        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With ERROR (OnStart)  " & ex.Message, sFuncName)
            Call WriteToLogFile("(OnStart) " & ex.Message, sFuncName)
        End Try

    End Sub

    Protected Overrides Sub OnStop()
        ' Add code here to perform any tear-down necessary to stop your service.

        Try
            p_iDebugMode = DEBUG_ON

            sFuncName = "Onstop()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Email Trigger Service Stops  " & Format(Now.Date, "dd-MMM-yyyy"), sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("----------------------------------------------------------------------", sFuncName)
            oEmailTrigger.Stop()

        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With ERROR (OnStop)  " & ex.Message, sFuncName)
            Call WriteToLogFile("(OnStop) " & ex.Message, sFuncName)
        End Try


    End Sub


    Private Sub EmailNotifiation()

        ' **********************************************************************************
        '   Function   :    EmailNotifiation()
        '   Purpose    :    This function Trigger Emails if any PO /PR send for an approval
        '
        '   Author     :    JOHN
        '   Date       :    07 April 2015
        ' **********************************************************************************

        Dim sFuncName As String = String.Empty
        Dim sSQL As String = String.Empty
        Dim oDT_EmailStatus As DataTable = Nothing
        Dim oDT_ApprovalMAtrix As DataTable = Nothing
        Dim oDV_ApprovalMatrix As DataView = Nothing
        Dim p_SyncDateTime As String = String.Empty
        Dim sBodyH As String = String.Empty
        Dim sBodyL As String = String.Empty


        p_iDebugMode = DEBUG_ON
        Try
            sFuncName = "EmailNotifiation()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function " & Now.ToLongDateString, sFuncName)

            ''sSQL = "update " & p_oCompDef.sSAPDBName & ".. [AB_EmailStatus] set Status = 'Closed' where Fcount >= 3 "
            ''If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Status Close " & sSQL, sFuncName)

            ''If sSQL.Length >= 1 Then
            ''    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery() ", sFuncName)
            ''    ExecuteSQLInsertQuery(sSQL)
            ''End If


            sSQL = "SELECT TOP 50 [Sno],[DocType],[ObjectType],[Entity],[EmailID],[EmailBody],[EmailSub],[Status],  RIGHT( T0.EmailSub, LEN( T0.EmailSub) - (CHARINDEX('Draft No.', T0.EmailSub) + 9)) [Dockey] FROM " & p_oCompDef.sSAPDBName & ".. [AB_EmailStatus] T0 " & _
                "WHERE T0.[Status] = 'Open'"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Email Status " & sSQL, sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT() ", sFuncName)
            oDT_EmailStatus = ExecuteSQLQuery_DT(sSQL)
            If oDT_EmailStatus.Rows.Count = 0 Then Exit Sub

            sSQL = String.Empty
            For imjs As Integer = 0 To oDT_EmailStatus.Rows.Count - 1

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling SendEmailNotification() ", sFuncName)
                If SendEmailNotification(oDT_EmailStatus.Rows(imjs).Item("EmailBody").ToString.Trim, _
                                         oDT_EmailStatus.Rows(imjs).Item("EmailSub").ToString.Trim, oDT_EmailStatus.Rows(imjs).Item("EmailID").ToString.Trim _
                                         , sErrDesc) <> RTN_SUCCESS Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                    Call WriteToLogFile(sErrDesc, sFuncName)
                    '' sSQL += "UPDATE " & p_oCompDef.sSAPDBName & ".. [AB_EmailStatus] SET [ErrMsg] = '" & sErrDesc & "', [EmailDate] = " & Now.Date & ", [EmailTime] = '" & Now.ToShortTimeString & "', [Fcount] = isnull([Fcount],0) + 1 WHERE [Sno] = '" & oDT_EmailStatus.Rows(imjs).Item("Sno").ToString.Trim & "'"
                    sSQL += "UPDATE " & p_oCompDef.sSAPDBName & ".. [AB_EmailStatus] SET [Status] = 'Fail', [ErrMsg] = '" & Replace(sErrDesc, "'", "''") & "', [EmailDate] = DATEADD(day,datediff(day,0,GETDATE()),0), [EmailTime] = '" & Now.ToShortTimeString & "' WHERE [Sno] = '" & oDT_EmailStatus.Rows(imjs).Item("Sno").ToString.Trim & "'"

                    sBodyL = sBodyL & " " & " Draft Key   : " & oDT_EmailStatus.Rows(imjs).Item("Dockey").ToString.Trim & " .<br />"
                    sBodyL = sBodyL & " " & " Object Type : " & oDT_EmailStatus.Rows(imjs).Item("DocType").ToString.Trim & " .<br />"
                    sBodyL = sBodyL & " " & " Entity      : " & oDT_EmailStatus.Rows(imjs).Item("Entity").ToString.Trim & " .<br />"
                    sBodyL = sBodyL & " " & " Error Msg   : " & sErrDesc
                    sBodyL = sBodyL & "<br /><br />"
                    sBodyL = sBodyL & " Please do not reply to this email. <div/>"
                Else
                    sSQL += "UPDATE " & p_oCompDef.sSAPDBName & ".. [AB_EmailStatus] SET [Status] = 'Closed',  [EmailDate] = DATEADD(day,datediff(day,0,GETDATE()),0) , [EmailTime] = '" & Now.ToShortTimeString & "' WHERE [Sno] = '" & oDT_EmailStatus.Rows(imjs).Item("Sno").ToString.Trim & "'"
                End If
            Next imjs
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Update Sql " & sSQL, sFuncName)

            If sSQL.Length > 1 Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery() ", sFuncName)
                ExecuteSQLInsertQuery(sSQL)
            End If

            If sBodyL.Length > 1 Then
                p_SyncDateTime = Format(Now, "dddd") & ", " & Format(Now, "MMM") & " " & Format(Now, "dd") & ", " & Format(Now, "yyyy") & " " & Format(Now, "HH:mm:ss")
                sBodyH = sBodyH & "<div align=left style='font-size:10.0pt;font-family:Arial'>"
                sBodyH = sBodyH & " Dear Sir/Madam,<br /><br />"
                sBodyH = sBodyH & p_SyncDateTime & " <br /><br />"
                sBodyH = sBodyH & "Approval Email notifications are failed for the below Draft key and rectify for Email trigger <br /><br />"

                sBodyL = sBodyL & "<br /><br /> Please do not reply to this email. <div/>"

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling SendEmailNotification() 2 ", sFuncName)
                If SendEmailNotification(sBodyH & sBodyL & "", _
                                         "Email Error Nofication", p_oCompDef.sToEmailID _
                                         , sErrDesc) <> RTN_SUCCESS Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                    Call WriteToLogFile(sErrDesc, sFuncName)

                End If
                
            End If


        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
            Call WriteToLogFile(sErrDesc, sFuncName)
        End Try
    End Sub

    Private Sub EmailNotifiation_OLD()

        ' **********************************************************************************
        '   Function   :    EmailNotifiation()
        '   Purpose    :    This function Trigger Emails if any PO send for an approval
        '
        '   Author     :    JOHN
        '   Date       :    07 April 2015
        ' **********************************************************************************

        Dim sFuncName As String = String.Empty
        Dim sSQL As String = String.Empty
        Dim oDT_ODRF As DataTable = Nothing
        Dim oDT_ApprovalMAtrix As DataTable = Nothing
        Dim oDV_ApprovalMatrix As DataView = Nothing
        p_iDebugMode = DEBUG_ON
        Try
            sFuncName = "EmailNotifiation()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function " & Now.ToLongDateString, sFuncName)
            sSQL = "SELECT TOP (50) T0.[USER_CODE], T0.[U_NAME],DocEntry, U_AB_APPROVALAMT, DocEntry FROM OUSR T0  INNER JOIN ODRF T1 ON T0.[USERID] = T1.[UserSign] " & _
                "WHERE isnull(T1.[U_AB_EmailFlag],'N') = 'N' and T1.ObjType = '22'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("List of PO Draft " & sSQL, sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT() ", sFuncName)
            oDT_ODRF = ExecuteSQLQuery_DT(sSQL)
            If oDT_ODRF.Rows.Count = 0 Then Exit Sub

            sSQL = "SELECT T0.[U_AB_APPROVER1_U],  T1.E_Mail [App1], T0.[U_AB_APPROVER2_U], T2.E_Mail [App2], T0.[U_AB_APPROVER3_U] , T0.[U_AB_REQUESTOR]," & _
             "T3.E_Mail [App3], isnull(T1.E_Mail,'') + ',' + isnull(T2.E_Mail,'')  + ',' + isnull(T3.E_Mail,'') [Sender], T0.[U_AB_AMOUNTFROM] [AmountFrom], T0.[U_AB_AMOUNTTO] [AmountTo]  FROM [dbo].[@AE_APPROVALMATRIX]  T0 " & _
            "left outer join OUSR T1 on T1.USER_CODE = T0.U_AB_APPROVER1_U " & _
            "left outer join OUSR T2 on T2.USER_CODE = T0.U_AB_APPROVER1_U " & _
            "left outer join OUSR T3 on T3.USER_CODE = T0.U_AB_APPROVER1_U " & _
             "WHERE T0.[U_AB_DOCTYPE] = 'PO'"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Approval Matrix Query " & sSQL, sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT() ", sFuncName)
            oDT_ApprovalMAtrix = ExecuteSQLQuery_DT(sSQL)

            oDV_ApprovalMatrix = New DataView(oDT_ApprovalMAtrix)
            sSQL = String.Empty
            For imjs As Integer = 0 To oDT_ODRF.Rows.Count - 1

                oDV_ApprovalMatrix.RowFilter = "U_AB_REQUESTOR = '" & oDT_ODRF.Rows(imjs).Item("USER_CODE").ToString & "'" & _
                    "and AmountFrom >= '" & CDbl(oDT_ODRF.Rows(imjs).Item("U_AB_APPROVALAMT").ToString) & "' and AmountTo <= '" & CDbl(oDT_ODRF.Rows(imjs).Item("U_AB_APPROVALAMT").ToString) & "'"

                If oDV_ApprovalMatrix.Count = 0 Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No Rows were selected ", sFuncName)
                    Exit For
                End If


                'If SendEmailNotification(p_oCompDef.sSAPDBName, oDV_ApprovalMatrix.Item(0)("Sender").ToString, sErrDesc) <> RTN_SUCCESS Then
                '    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                '    Call WriteToLogFile(sErrDesc, sFuncName)
                '    Exit For
                'End If

                sSQL += "Update ODRF set U_AB_EmailFlag = 'Y' where DocEntry = '" & oDT_ODRF.Rows(imjs).Item("DocEntry").ToString & "'"


            Next imjs

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Update Sql " & sSQL, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT() ", sFuncName)

            oDT_ApprovalMAtrix = ExecuteSQLQuery_DT(sSQL)


        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
            Call WriteToLogFile(sErrDesc, sFuncName)
        End Try
    End Sub
End Class

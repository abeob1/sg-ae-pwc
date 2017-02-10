Namespace AE_PWC_AO06
    Module modItemList

        Public Sub ItemList_SBO_ItemEvent(ByVal FormUID As String, ByVal pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean, ByVal objForm As SAPbouiCOM.Form)
            Dim sFuncName As String = "ApprovalWindow_SBO_ItemEvent"
            Dim sErrDesc As String = String.Empty

            Try
                If pval.Before_Action = True Then
                    Select Case pval.EventType
                        Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                            objForm = p_oSBOApplication.Forms.GetForm(pval.FormTypeEx, pval.FormTypeCount)
                            If pval.ItemUID = "12" Then
                                Dim oForm_AprlWind As SAPbouiCOM.Form
                                oForm_AprlWind = p_oSBOApplication.Forms.GetFormByTypeAndCount("APRL", ItemListCount)
                                Dim oEdit As SAPbouiCOM.EditText = oForm_AprlWind.Items.Item(sItem).Specific
                            End If

                    End Select
                Else
                    Select Case pval.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE
                            objForm = p_oSBOApplication.Forms.GetForm(pval.FormTypeEx, pval.FormTypeCount)
                            Try
                                Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.Item("ITL")
                                oForm.Visible = True
                                ' BubbleEvent = False
                                Exit Try
                            Catch ex As Exception
                                p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Try
                            End Try
                            Exit Sub

                    End Select
                End If
            Catch ex As Exception
                sErrDesc = ex.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                Throw New ArgumentException(sErrDesc)
            End Try
        End Sub

    End Module
End Namespace


Public Class oGRPO
#Region "Build Table Structure"
    Private Function BuildTableOPDN() As DataTable
        Dim dt As New DataTable("OPDN")
        dt.Columns.Add("U_POSTxNo")
        dt.Columns.Add("CardCode")
        dt.Columns.Add("CardName")
        dt.Columns.Add("DocDate")
        Return dt
    End Function
    Private Function BuildTablePDN1() As DataTable
        Dim dt As New DataTable("PDN1")
        dt.Columns.Add("LineNum")
        dt.Columns.Add("ItemCode")
        dt.Columns.Add("Dscription")
        dt.Columns.Add("WhsCode")
        dt.Columns.Add("Quantity")
        dt.Columns.Add("VatGroup")
        dt.Columns.Add("PriceBefDi")
        dt.Columns.Add("BaseLine")
        dt.Columns.Add("BaseEntry")
        dt.Columns.Add("BaseType")
        Return dt
    End Function
    Private Function BuildTableSRNT() As DataTable
        Dim dt As New DataTable("SRNT")
        dt.Columns.Add("DistNumber")
        dt.Columns.Add("DocLineNum")
        Return dt
    End Function
#End Region

#Region "Insert into Table"
    Private Function InsertIntoOPDN(dt As DataTable, dr As DataRow) As DataTable
        Dim drNew As DataRow = dt.NewRow
        drNew("U_POSTxNo") = dr("U_POSTxNo")
        drNew("CardCode") = dr("CardCode")
        drNew("CardName") = dr("CardName")
        drNew("DocDate") = CDate(dr("DocDate")).ToString("yyyyMMdd")
        dt.Rows.Add(drNew)
        Return dt
    End Function
    Private Function InsertIntoPDN1(dt As DataTable, dr As DataRow) As DataTable
        Dim drNew As DataRow = dt.NewRow
        drNew("ItemCode") = dr("ItemCode")
        drNew("Dscription") = dr("Dscription")
        drNew("WhsCode") = dr("WhsCode")
        drNew("Quantity") = dr("Quantity")

        'drNew("VatGroup") = dr("VatGroup")
        'drNew("PriceBefDi") = dr("PriceBefDi")
        drNew("BaseLine") = dr("BaseLine")
        drNew("BaseEntry") = dr("BaseEntry")
        drNew("BaseType") = "22"

        drNew("LineNum") = dr("BaseLine") 'baseline can not be duplicate(if can, we don't need to sum up)
        dt.Rows.Add(drNew)
        Return dt
    End Function
    Private Function InsertIntoSRNT(dt As DataTable, dr As DataRow) As DataTable
        Dim drNew As DataRow = dt.NewRow
        drNew("DistNumber") = dr("DistNumber")
        drNew("DocLineNum") = dr("BaseLine")
        dt.Rows.Add(drNew)
        Return dt
    End Function
#End Region
    Public Sub CreateGRPO()
        Try
            Dim DocType As String = "20"
            Dim cn As New Connection
            Dim xm As New oXML
            Dim oRunning As New oRunningMonitor
            Dim dt As DataTable = cn.Integration_RunQuery("sp_GRPO_LoadForSync")
            If Not IsNothing(dt) Then
                If PublicVariable.oCompanyInfo.Connected = False Then
                    xm.SetDB()
                    Dim sErrMsg As String = xm.ConnectSAPDB()
                    If sErrMsg <> "" Then
                        Functions.WriteLog(sErrMsg)
                        Return
                    End If
                End If
                If PublicVariable.oCompanyInfo.InTransaction Then
                    PublicVariable.oCompanyInfo.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                End If
                For Each dr As DataRow In dt.Rows
                    Dim HeaderID As Integer = dr.Item("ID")
                    oRunning.UpdateMonitor("GRPO", HeaderID)
                    Dim ret As String = ""
                    Dim ds As New DataSet
                    Dim dtOWTR As DataTable = BuildTableOPDN()
                    Dim dtWTR1 As DataTable = BuildTablePDN1()
                    Dim dtSRNT As DataTable = BuildTableSRNT()
                    '----------add header----------
                    dtOWTR = InsertIntoOPDN(dtOWTR, dr)

                    '----------add line------------
                    Dim dtLine As DataTable = cn.Integration_RunQuery("sp_GRPOLine_LoadByID " + CStr(HeaderID))
                    If dtLine.Rows.Count = 0 Then
                        ret = "Service Return: GRPO has no line item."
                        cn.Integration_RunQuery("sp_GRPO_UpdateReceived " + CStr(HeaderID) + ",'" + ret + "'")
                        Continue For
                    Else
                        For Each drLine As DataRow In dtLine.Rows
                            dtWTR1 = InsertIntoPDN1(dtWTR1, drLine)
                        Next
                    End If

                    '----------add serial----------
                    Dim dtSerial As DataTable = cn.Integration_RunQuery("sp_SerialNumber_LoadByID " + CStr(HeaderID) + ",'" + DocType + "'")
                    For Each drSerial As DataRow In dtSerial.Rows
                        dtSRNT = InsertIntoSRNT(dtSRNT, drSerial)
                    Next

                    ds.Tables.Add(dtOWTR.Copy)
                    ds.Tables.Add(dtWTR1.Copy)
                    ds.Tables.Add(dtSRNT.Copy)


                    Dim xmlstr As String = xm.ToXMLStringFromDS(DocType, ds)

                    ret = xm.CreateMarketingDocument(xmlstr, DocType)
                    If xmlstr.Contains("'") Then
                        xmlstr = xmlstr.Replace("'", " ")
                    End If
                    If ret.Contains("'") Then
                        ret = ret.Replace("'", " ")
                    End If
                    Functions.WriteXMLLog(DocType, xmlstr, ret)
                   
                    cn.Integration_RunQuery("sp_GRPO_UpdateReceived " + CStr(HeaderID) + ",'" + ret + "'")
                Next
            End If
        Catch ex As Exception
            Functions.WriteLog("CreateGRPO:" + ex.ToString)
        End Try
    End Sub
End Class

Public Class oReturn
#Region "Build Table Structure"
    Private Function BuildTableORPD() As DataTable
        Dim dt As New DataTable("ORPD")
        dt.Columns.Add("U_POSTxNo")
        dt.Columns.Add("CardCode")
        dt.Columns.Add("CardName")
        dt.Columns.Add("DocDate")
        dt.Columns.Add("DocDueDate")
        Return dt
    End Function
    Private Function BuildTableRPD1() As DataTable
        Dim dt As New DataTable("RPD1")
        dt.Columns.Add("LineNum")
        dt.Columns.Add("ItemCode")
        dt.Columns.Add("Dscription")
        dt.Columns.Add("WhsCode")
        dt.Columns.Add("Quantity")
        dt.Columns.Add("VatGroup")
        dt.Columns.Add("LineTotal")
        dt.Columns.Add("U_RMA")
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
    Private Function InsertIntoORPD(dt As DataTable, dr As DataRow) As DataTable
        Dim drNew As DataRow = dt.NewRow
        drNew("U_POSTxNo") = dr("U_POSTxNo")
        drNew("CardCode") = dr("CardCode")
        drNew("CardName") = dr("CardName")
        drNew("DocDate") = CDate(dr("DocDate")).ToString("yyyyMMdd")
        drNew("DocDueDate") = CDate(dr("DocDueDate")).ToString("yyyyMMdd")
        dt.Rows.Add(drNew)
        Return dt
    End Function
    Private Function InsertIntoRPD1(dt As DataTable, dr As DataRow) As DataTable
        Dim drNew As DataRow = dt.NewRow
        drNew("LineNum") = dr("LineNum")
        drNew("ItemCode") = dr("ItemCode")
        drNew("Dscription") = dr("Dscription")
        drNew("WhsCode") = dr("WhsCode")
        drNew("Quantity") = dr("Quantity")
        drNew("VatGroup") = dr("VatGroup")
        drNew("U_RMA") = dr("U_RMA").ToString
        drNew("LineTotal") = GetLastPurchasePrice(dr("ItemCode").ToString, dr("WhsCode")) * dr("Quantity")
        dt.Rows.Add(drNew)
        Return dt
    End Function
    Private Function InsertIntoSRNT(dt As DataTable, dr As DataRow) As DataTable
        Dim drNew As DataRow = dt.NewRow
        drNew("DistNumber") = dr("DistNumber")
        drNew("DocLineNum") = dr("DocLineNum")
        dt.Rows.Add(drNew)
        Return dt
    End Function
#End Region
    Public Sub CreateGoodsReturn()
        Try
            Dim xm As New oXML
            Dim DocType As String = "21"
            Dim cn As New Connection
            Dim oRunning As New oRunningMonitor
            Dim dt As DataTable = cn.Integration_RunQuery("sp_GoodsReturn_LoadForSync")
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
                    Dim HeaderID As String = dr.Item("ID")
                    oRunning.UpdateMonitor("Return", HeaderID)
                    Dim ret As String = ""
                    Dim ds As New DataSet
                    Dim dtOWTR As DataTable = BuildTableORPD()
                    Dim dtWTR1 As DataTable = BuildTableRPD1()
                    Dim dtSRNT As DataTable = BuildTableSRNT()
                    '----------add header----------
                    dtOWTR = InsertIntoORPD(dtOWTR, dr)

                    '----------add line------------
                    Dim dtLine As DataTable = cn.Integration_RunQuery("sp_GoodsReturnLine_LoadByID " + CStr(HeaderID))
                    If dtLine.Rows.Count = 0 Then
                        ret = "Service Return: GoodsReturn has no line item."
                        cn.Integration_RunQuery("sp_GoodsReturn_UpdateReceived " + CStr(HeaderID) + ",'" + ret + "'")
                        Continue For
                    Else
                        For Each drLine As DataRow In dtLine.Rows
                            dtWTR1 = InsertIntoRPD1(dtWTR1, drLine)
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
                    If ret.Contains("'") Then
                        ret = ret.Replace("'", " ")
                    End If
                    If xmlstr.Contains("'") Then
                        xmlstr = xmlstr.Replace("'", " ")
                    End If
                    Functions.WriteXMLLog(DocType, xmlstr, ret)
                    
                    cn.Integration_RunQuery("sp_GoodsReturn_UpdateReceived " + CStr(HeaderID) + ",'" + ret + "'")
                Next
            End If
        Catch ex As Exception
            Functions.WriteLog(ex.ToString)
        End Try
    End Sub

    Private Function GetLastPurchasePrice(ItemCode As String, WhsCode As String) As Double
        Dim cn As New Connection
        Dim dt As DataTable = cn.SAP_RunQuery("exec sp_AI_GetLastPurPrice_ByWhs '" + WhsCode + "','" + ItemCode + "'")
        If dt.Rows.Count > 0 Then
            Return dt.Rows(0).Item("PriceBefDi")
        Else
            Return 0
        End If
    End Function
End Class

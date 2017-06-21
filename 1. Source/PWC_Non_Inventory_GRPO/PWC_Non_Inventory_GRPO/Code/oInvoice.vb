Public Class oInvoice
    Dim DocEntry_Invoice As Integer = 0
#Region "Build Table Structure"
    Private Function BuildTableOINV() As DataTable ' Invoice
        Dim dt As New DataTable("OINV")
        dt.Columns.Add("U_POSTxNo")
        dt.Columns.Add("CardCode")
        dt.Columns.Add("DocDate")
        dt.Columns.Add("TaxDate")
        dt.Columns.Add("Posted")
        Return dt
    End Function
    Private Function BuildTableINV1() As DataTable 'Invoice detail
        Dim dt As New DataTable("INV1")
        dt.Columns.Add("LineNum")
        dt.Columns.Add("ItemCode")
        dt.Columns.Add("Dscription")
        dt.Columns.Add("WhsCode")
        dt.Columns.Add("Quantity")
        dt.Columns.Add("VatGroup")
        dt.Columns.Add("PriceAfVAT")
        dt.Columns.Add("OcrCode")
        dt.Columns.Add("U_PromoCode")
        dt.Columns.Add("U_PromoMPDT")
        dt.Columns.Add("SlpCode")
        Return dt
    End Function
    Private Function BuildTableINV9() As DataTable 'DownPayment
        Dim dt As New DataTable("INV9")
        dt.Columns.Add("BaseAbs") 'Downpayment Invoice DocEntry
        dt.Columns.Add("DocEntry") ' AR Invoice DocEntry
        dt.Columns.Add("DrawnSum")
        Return dt
    End Function
    Private Function BuildTableINV11() As DataTable 'DownPayment Detail
        Dim dt As New DataTable("INV11")
        dt.Columns.Add("ItemCode")
        dt.Columns.Add("Dscription")
        Return dt
    End Function
    Private Function BuildTableSRNT() As DataTable 'Serial Number
        Dim dt As New DataTable("SRNT")
        dt.Columns.Add("DistNumber")
        dt.Columns.Add("DocLineNum")
        dt.Columns.Add("Notes")
        Return dt
    End Function
    Private Function BuildTableORCT() As DataTable 'Incoming Payment
        Dim dt As New DataTable("ORCT")
        dt.Columns.Add("DocDate")
        dt.Columns.Add("DocDueDate")
        dt.Columns.Add("CardCode")
        dt.Columns.Add("CashAcct")
        dt.Columns.Add("CashSum")
        dt.Columns.Add("TrsfrAcct")
        dt.Columns.Add("TrsfrSum")
        dt.Columns.Add("DocType")
        Return dt
    End Function
    Private Function BuildTableRCT2() As DataTable 'Incoming Payment Detail
        Dim dt As New DataTable("RCT2")
        dt.Columns.Add("DocEntry")
        dt.Columns.Add("InvType")
        dt.Columns.Add("SumApplied")
        dt.Columns.Add("DocLine")
        Return dt
    End Function
    Private Function BuildTableRCT3() As DataTable 'Incoming Payment Detail
        Dim dt As New DataTable("RCT3")
        dt.Columns.Add("CreditCard")
        dt.Columns.Add("CardValid")
        dt.Columns.Add("CrCardNum")
        dt.Columns.Add("CreditSum")
        dt.Columns.Add("VoucherNum")
        dt.Columns.Add("CreditAcct")
        Return dt
    End Function
#End Region

#Region "Insert into Table"
    Private Function InsertIntoOINV(dt As DataTable, dr As DataRow) As DataTable
        Dim drNew As DataRow = dt.NewRow
        drNew("U_POSTxNo") = dr("U_POSTxNo")
        drNew("CardCode") = dr("CardCode")
        drNew("DocDate") = CDate(dr("DocDate")).ToString("yyyyMMdd")
        drNew("TaxDate") = CDate(dr("DocDate")).ToString("yyyyMMdd")
        drNew("Posted") = "Y"
        dt.Rows.Add(drNew)
        Return dt
    End Function
    Private Function InsertIntoINV1(dt As DataTable, dr As DataRow) As DataTable
        Dim drNew As DataRow = dt.NewRow
        drNew("LineNum") = dr("LineNum")
        drNew("ItemCode") = dr("ItemCode")
        drNew("Dscription") = dr("Dscription")
        drNew("WhsCode") = dr("WhsCode")
        drNew("Quantity") = dr("Quantity")
        drNew("VatGroup") = ApplyGST(dr("GST"))
        drNew("PriceAfVAT") = dr("GrossPrice")
        drNew("OcrCode") = dr("CostCenter")
        drNew("U_PromoCode") = dr("PromoCode")
        drNew("U_PromoMPDT") = dr("PromoMandatoryProduct")
        Dim slpcode As Integer = GetSlpCodeByName(dr("SalesEmployeeName"))
        If slpcode > 0 Then
            drNew("SlpCode") = slpcode
        End If

        dt.Rows.Add(drNew)
        Return dt
    End Function
    Private Function InsertIntoINV9(dt As DataTable, dr As DataRow) As DataTable
        Dim DownPaymentNo As Integer = 0
        DownPaymentNo = GetInvoiceEntryByPOSNo(dr.Item("DownpaymentNo"), "203")

        Dim drNew As DataRow = dt.NewRow
        drNew("BaseAbs") = DownPaymentNo
        'drNew("DocEntry") = dr("Dscription")
        'drNew("DrawnSum") = dr("WhsCode")
        dt.Rows.Add(drNew)
        Return dt
    End Function
    Private Function InsertIntoSRNT(dt As DataTable, dr As DataRow) As DataTable
        Dim drNew As DataRow = dt.NewRow
        drNew("DistNumber") = dr("DistNumber")
        drNew("DocLineNum") = dr("DocLineNum")
        drNew("Notes") = dr("Notes")
        dt.Rows.Add(drNew)
        Return dt
    End Function
    Private Function InsertIntoRCT2(dt As DataTable, InvoiceType As String) As DataTable
        Dim drNew As DataRow = dt.NewRow
        drNew("DocEntry") = DocEntry_Invoice
        If InvoiceType = "RES" Then
            drNew("InvType") = "203"
        Else
            drNew("InvType") = "13"
        End If
        'drNew("SumApplied") = dr("Amount")
        drNew("DocLine") = 0
        dt.Rows.Add(drNew)
        Return dt
    End Function
    Private Function InsertIntoRCT3(dt As DataTable, dr As DataRow) As DataTable
        Dim drNew As DataRow = dt.NewRow
        drNew("CreditCard") = GetCrCCodeByName(dr("PaymentMethod"))
        drNew("CardValid") = "99990101"
        drNew("CrCardNum") = "999"
        drNew("CreditSum") = dr("Amount")
        drNew("VoucherNum") = "1"
        drNew("CreditAcct") = GetCreditCardGL(dr("PaymentMethod"))
        dt.Rows.Add(drNew)
        Return dt
    End Function
    Private Function InsertIntoORCT(dt As DataTable, dr As DataRow) As DataTable
        Dim drNew As DataRow = dt.NewRow
        drNew("DocDate") = CDate(dr("DocDate")).ToString("yyyyMMdd")
        drNew("DocDueDate") = CDate(dr("DocDate")).ToString("yyyyMMdd")
        drNew("CardCode") = dr("CardCode")

        drNew("CashAcct") = PublicVariable.pmCashAcct
        drNew("CashSum") = dr("CashAmount")
        drNew("TrsfrAcct") = PublicVariable.pmTransferAcct
        drNew("TrsfrSum") = dr("TransferAmount")

        drNew("DocType") = "C"

        dt.Rows.Add(drNew)
        Return dt
    End Function
#End Region

#Region "Functions and Mapping"
    Private Function GetInvoiceEntryByPOSNo(POSTxNo As String, DocType As String) As Integer
        Dim cn As New Connection
        Dim strQuery As String = ""
        If DocType = "13" Then
            strQuery = "select max(DocEntry) DocEntry from OINV where isnull(U_POSTxNo,'')='" + POSTxNo + "'"
        Else
            strQuery = "select max(DocEntry) DocEntry from ODPI where isnull(U_POSTxNo,'')='" + POSTxNo + "'"
        End If
        Dim dt As DataTable = cn.SAP_RunQuery(strQuery)
        If dt.Rows.Count > 0 Then
            Return dt.Rows(0).Item("DocEntry")
        Else
            Return 0
        End If
    End Function
    Private Function ApplyGST(Apply As String) As String
        If Apply = "N" Then
            Return PublicVariable.NonGSTCode
        Else
            Return PublicVariable.GSTCode
        End If
    End Function
    Private Function GetSlpCodeByName(slpName As String) As Integer
        Dim slpCode As Integer = -1
        Dim strQuery As String = ""
        Dim cn As New Connection
        strQuery = "select slpcode  from OSLP where isnull(slpName,'')='" + slpName + "'"
        Dim dt As DataTable = cn.SAP_RunQuery(strQuery)
        If dt.Rows.Count > 0 Then
            Return dt.Rows(0).Item("slpcode")
        Else
            Return -1
        End If
        Return slpCode
    End Function
    Private Function GetCrCCodeByName(CrCName As String) As Integer
        Dim CrCCode As Integer = 0
        Dim strQuery As String = ""
        Dim cn As New Connection
        strQuery = "select CreditCard from OCRC where isnull(CardName,'')='" + CrCName + "'"
        Dim dt As DataTable = cn.SAP_RunQuery(strQuery)
        If dt.Rows.Count > 0 Then
            Return dt.Rows(0).Item("CreditCard")
        Else
            Return 0
        End If
        Return CrCCode
    End Function
    Private Function GetCreditCardGL(CrCName As String) As Integer
        Dim CrCCode As Integer = 0
        Dim strQuery As String = ""
        Dim cn As New Connection
        strQuery = "select AcctCode from OCRC where isnull(CardName,'')='" + CrCName + "'"
        Dim dt As DataTable = cn.SAP_RunQuery(strQuery)
        If dt.Rows.Count > 0 Then
            Return dt.Rows(0).Item("AcctCode")
        Else
            Return 0
        End If
        Return CrCCode
    End Function
#End Region
    
#Region "Create Invoice"
    Public Sub CreateInvoice()
        Dim DocType As String = "13"
        Dim cn As New Connection
        Dim xm As New oXML
        Dim oRunning As New oRunningMonitor
        Try
            Dim dt As DataTable = cn.Integration_RunQuery("exec sp_Invoice_LoadForSync")
            If Not IsNothing(dt) Then
                'If PublicVariable.oCompanyInfo.Connected = False Then
                xm.SetDB()
                Dim sErrMsg As String = xm.ConnectSAPDB()
                If sErrMsg <> "" Then
                    Functions.WriteLog(sErrMsg)
                    Return
                End If
                'End If
                If PublicVariable.oCompanyInfo.InTransaction Then
                    PublicVariable.oCompanyInfo.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                End If
                For Each dr As DataRow In dt.Rows
                    DocType = "13"
                    Dim HeaderID As String = dr.Item("ID")
                    oRunning.UpdateMonitor("Invoice", HeaderID)
                    Dim ret As String = ""
                    Dim ds As New DataSet
                    Dim dtOINV As DataTable = BuildTableOINV()
                    Dim dtINV1 As DataTable = BuildTableINV1()
                    Dim dtSRNT As DataTable = BuildTableSRNT()
                    Dim dtINV9 As DataTable = BuildTableINV9()

                    Dim Flag As String = dr("Flag").ToString
                    If InvoicePaymentValidation(HeaderID) = False Then
                        If Flag <> "1" Then
                            ret = "Service Return: Invoice Amount and Payment Amount does not matching."
                            cn.Integration_RunQuery("sp_Invoice_UpdateReceived '" + CStr(HeaderID) + "','" + ret + "'")
                            Continue For
                        End If
                    End If

                    '----------add Invoice header----------
                    dtOINV = InsertIntoOINV(dtOINV, dr)

                    '----------add Invoice line------------
                    Dim dtLine As DataTable = cn.Integration_RunQuery("sp_InvoiceLine_LoadByID " + CStr(HeaderID))
                    If dtLine.Rows.Count = 0 Then
                        ret = "Service Return: Invoice has no line item."
                        cn.Integration_RunQuery("sp_Invoice_UpdateReceived '" + CStr(HeaderID) + "','" + ret + "'")
                        Continue For
                    Else
                        For Each drLine As DataRow In dtLine.Rows
                            dtINV1 = InsertIntoINV1(dtINV1, drLine)
                        Next
                    End If

                    '----------add serial----------
                    Dim dtSerial As DataTable = cn.Integration_RunQuery("sp_SerialNumber_LoadByID " + CStr(HeaderID) + ",'13'")
                    For Each drSerial As DataRow In dtSerial.Rows
                        dtSRNT = InsertIntoSRNT(dtSRNT, drSerial)
                    Next

                    Dim InvoiceType As String = dr("InvoiceType").ToString
                    If InvoiceType = "RES" Then 'Downpayment for Reservation
                        dtOINV.TableName = "ODPI"
                        dtINV1.TableName = "DPI1"
                        DocType = "203"
                    End If
                    ds.Tables.Add(dtOINV.Copy)
                    ds.Tables.Add(dtINV1.Copy)
                    ds.Tables.Add(dtSRNT.Copy)

                    'Reservation - downpayment invoice
                    If dr.Item("DownpaymentNo").ToString <> "" Then

                        InsertIntoINV9(dtINV9, dr)
                        ds.Tables.Add(dtINV9.Copy)
                    End If
                    If PublicVariable.oCompanyInfo.InTransaction Then
                        PublicVariable.oCompanyInfo.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    End If

                    PublicVariable.oCompanyInfo.StartTransaction()


                    Dim xmlstr As String = xm.ToXMLStringFromDS(DocType, ds)
                    'CREATE INVOICE
                    ret = xm.CreateMarketingDocument(xmlstr, DocType)
                    If xmlstr.Contains("'") Then
                        xmlstr = xmlstr.Replace("'", " ")
                    End If
                    If ret.Contains("'") Then
                        ret = ret.Replace("'", " ")
                    End If
                    Functions.WriteXMLLog(DocType, xmlstr, ret)
                    If ret = "" Then
                        DocEntry_Invoice = PublicVariable.oCompanyInfo.GetNewObjectKey()

                        If dr("PaymentType").ToString <> "BLK" Then
                            Dim dtPayment As DataTable = cn.Integration_RunQuery("sp_PaymentMean_LoadByID " + CStr(HeaderID))
                            If dtPayment.Rows.Count = 0 Then
                                ret = "Service Return: Payment Detail has no line item."
                                cn.Integration_RunQuery("sp_Invoice_UpdateReceived '" + CStr(HeaderID) + "','" + ret + "'")
                                Return
                            End If

                            Dim dtORCT As DataTable = BuildTableORCT()
                            Dim dtRCT2 As DataTable = BuildTableRCT2()
                            Dim dtRCT3 As DataTable = BuildTableRCT3()

                            If dr("PaymentType") = "IN" Then 'Incoming Payment
                                DocType = "24"
                            Else    'Outgoing Payment
                                dtORCT.TableName = "OVPM"
                                dtRCT2.TableName = "VPM2"
                                dtRCT3.TableName = "VPM3"
                                DocType = "46"
                            End If
                            'If dr("InvoiceType").ToString = "INV" Then
                            '    DocEntry_Invoice = GetInvoiceEntryByPOSNo(dr("U_POSTxNo").ToString, "13")
                            'Else
                            '    DocEntry_Invoice = GetInvoiceEntryByPOSNo(dr("U_POSTxNo").ToString, "203")
                            'End If


                            '----------add payment header: include cash and transfer----------
                            dtORCT = InsertIntoORCT(dtORCT, dtPayment.Rows(0))

                            '----------add payment invoice------------
                            dtRCT2 = InsertIntoRCT2(dtRCT2, InvoiceType)

                            '----------add payment credit card and others--------------

                            For Each drPayment As DataRow In dtPayment.Rows
                                If drPayment("Amount") <> 0 Then
                                    dtRCT3 = InsertIntoRCT3(dtRCT3, drPayment)
                                End If
                            Next

                            ds = New DataSet
                            ds.Tables.Add(dtORCT.Copy)
                            ds.Tables.Add(dtRCT2.Copy)
                            ds.Tables.Add(dtRCT3.Copy)

                            xmlstr = xm.ToXMLStringFromDS(DocType, ds)
                            ret = xm.CreateMarketingDocument(xmlstr, DocType)
                            If xmlstr.Contains("'") Then
                                xmlstr = xmlstr.Replace("'", " ")
                            End If
                            If ret.Contains("'") Then
                                ret = ret.Replace("'", " ")
                            End If
                            Functions.WriteXMLLog(DocType, xmlstr, ret)
                        End If

                        If ret = "" Then
                            If PublicVariable.oCompanyInfo.InTransaction Then
                                PublicVariable.oCompanyInfo.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                            End If
                        Else
                            If PublicVariable.oCompanyInfo.InTransaction Then
                                PublicVariable.oCompanyInfo.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            End If
                        End If

                        cn.Integration_RunQuery("sp_Invoice_UpdateReceived '" + CStr(HeaderID) + "','" + ret + "'")
                    Else
                        If PublicVariable.oCompanyInfo.InTransaction Then
                            PublicVariable.oCompanyInfo.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        End If
                        If ret.Contains("'") Then
                            ret = ret.Replace(",", "''")
                        End If
                        cn.Integration_RunQuery("sp_Invoice_UpdateReceived '" + CStr(HeaderID) + "','" + ret + "'")
                    End If
                Next
            End If
        Catch ex As Exception
            If PublicVariable.oCompanyInfo.InTransaction Then
                PublicVariable.oCompanyInfo.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
            Functions.WriteLog(ex.ToString)
        End Try
    End Sub

    Private Function InvoicePaymentValidation(InvoiceHeaderID As String) As Boolean
        Dim cn As New Connection
        Dim dt As DataTable = cn.Integration_RunQuery("sp_Invoice_Validation " + CStr(InvoiceHeaderID))
        If dt.Rows.Count > 0 Then
            If Double.Parse(dt.Rows(0)("PaymentAmt").ToString) <> Double.Parse(dt.Rows(0)("InvoiceAmt")) Then
                Return False
            End If
        End If
        Return True
    End Function

    Public Sub ManualCreateInvoice(HeaderID As String)
        Dim DocType As String = "13"
        Dim cn As New Connection
        Dim xm As New oXML
        Dim oRunning As New oRunningMonitor
        Try
            Dim dt As DataRow() = cn.Integration_RunQuery("sp_Invoice_LoadForSync").Select("ID=" + HeaderID)
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
                For Each dr As DataRow In dt
                    DocType = "13"
                    oRunning.UpdateMonitor("Invoice", HeaderID)
                    Dim ret As String = ""
                    Dim ds As New DataSet
                    Dim dtOINV As DataTable = BuildTableOINV()
                    Dim dtINV1 As DataTable = BuildTableINV1()
                    Dim dtSRNT As DataTable = BuildTableSRNT()
                    Dim dtINV9 As DataTable = BuildTableINV9()

                    If InvoicePaymentValidation(HeaderID) = False Then
                        ret = "Service Return: Invoice Amount and Payment Amount does not matching."
                        cn.Integration_RunQuery("sp_Invoice_UpdateReceived '" + CStr(HeaderID) + "','" + ret + "'")
                    Else

                        '----------add Invoice header----------
                        dtOINV = InsertIntoOINV(dtOINV, dr)

                        '----------add Invoice line------------
                        Dim dtLine As DataTable = cn.Integration_RunQuery("sp_InvoiceLine_LoadByID " + CStr(HeaderID))
                        If dtLine.Rows.Count = 0 Then
                            ret = "Service Return: Invoice has no line item."
                            cn.Integration_RunQuery("sp_Invoice_UpdateReceived '" + CStr(HeaderID) + "','" + ret + "'")
                            Return
                        Else
                            For Each drLine As DataRow In dtLine.Rows
                                dtINV1 = InsertIntoINV1(dtINV1, drLine)
                            Next
                        End If

                        '----------add serial----------
                        Dim dtSerial As DataTable = cn.Integration_RunQuery("sp_SerialNumber_LoadByID " + CStr(HeaderID) + ",'13'")
                        For Each drSerial As DataRow In dtSerial.Rows
                            dtSRNT = InsertIntoSRNT(dtSRNT, drSerial)
                        Next

                        Dim InvoiceType As String = dr("InvoiceType").ToString
                        If InvoiceType = "RES" Then 'Downpayment for Reservation
                            dtOINV.TableName = "ODPI"
                            dtINV1.TableName = "DPI1"
                            DocType = "203"
                        End If
                        ds.Tables.Add(dtOINV.Copy)
                        ds.Tables.Add(dtINV1.Copy)
                        ds.Tables.Add(dtSRNT.Copy)

                        'Reservation - downpayment invoice
                        If dr.Item("DownpaymentNo").ToString <> "" Then

                            InsertIntoINV9(dtINV9, dr)
                            ds.Tables.Add(dtINV9.Copy)
                        End If
                        If PublicVariable.oCompanyInfo.InTransaction Then
                            PublicVariable.oCompanyInfo.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        End If

                        PublicVariable.oCompanyInfo.StartTransaction()


                        Dim xmlstr As String = xm.ToXMLStringFromDS(DocType, ds)
                        'CREATE INVOICE
                        ret = xm.CreateMarketingDocument(xmlstr, DocType)
                        If xmlstr.Contains("'") Then
                            xmlstr = xmlstr.Replace("'", " ")
                        End If
                        If ret.Contains("'") Then
                            ret = ret.Replace("'", " ")
                        End If
                        Functions.WriteXMLLog(DocType, xmlstr, ret)
                        If ret = "" Then
                            DocEntry_Invoice = PublicVariable.oCompanyInfo.GetNewObjectKey()

                            If dr("PaymentType").ToString <> "BLK" Then
                                Dim dtPayment As DataTable = cn.Integration_RunQuery("sp_PaymentMean_LoadByID " + CStr(HeaderID))
                                If dtPayment.Rows.Count = 0 Then
                                    ret = "Service Return: Payment Detail has no line item."
                                    cn.Integration_RunQuery("sp_Invoice_UpdateReceived '" + CStr(HeaderID) + "','" + ret + "'")
                                    Return
                                End If

                                Dim dtORCT As DataTable = BuildTableORCT()
                                Dim dtRCT2 As DataTable = BuildTableRCT2()
                                Dim dtRCT3 As DataTable = BuildTableRCT3()

                                If dr("PaymentType") = "IN" Then 'Incoming Payment
                                    DocType = "24"
                                Else    'Outgoing Payment
                                    dtORCT.TableName = "OVPM"
                                    dtRCT2.TableName = "VPM2"
                                    dtRCT3.TableName = "VPM3"
                                    DocType = "46"
                                End If
                                'If dr("InvoiceType").ToString = "INV" Then
                                '    DocEntry_Invoice = GetInvoiceEntryByPOSNo(dr("U_POSTxNo").ToString, "13")
                                'Else
                                '    DocEntry_Invoice = GetInvoiceEntryByPOSNo(dr("U_POSTxNo").ToString, "203")
                                'End If


                                '----------add payment header: include cash and transfer----------
                                dtORCT = InsertIntoORCT(dtORCT, dtPayment.Rows(0))

                                '----------add payment invoice------------
                                dtRCT2 = InsertIntoRCT2(dtRCT2, InvoiceType)

                                '----------add payment credit card and others--------------

                                For Each drPayment As DataRow In dtPayment.Rows
                                    If drPayment("Amount") <> 0 Then
                                        dtRCT3 = InsertIntoRCT3(dtRCT3, drPayment)
                                    End If
                                Next

                                ds = New DataSet
                                ds.Tables.Add(dtORCT.Copy)
                                ds.Tables.Add(dtRCT2.Copy)
                                ds.Tables.Add(dtRCT3.Copy)

                                xmlstr = xm.ToXMLStringFromDS(DocType, ds)
                                ret = xm.CreateMarketingDocument(xmlstr, DocType)
                                If xmlstr.Contains("'") Then
                                    xmlstr = xmlstr.Replace("'", " ")
                                End If
                                If ret.Contains("'") Then
                                    ret = ret.Replace("'", " ")
                                End If
                                Functions.WriteXMLLog(DocType, xmlstr, ret)
                            End If

                            If ret = "" Then
                                If PublicVariable.oCompanyInfo.InTransaction Then
                                    PublicVariable.oCompanyInfo.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                End If
                            Else
                                If PublicVariable.oCompanyInfo.InTransaction Then
                                    PublicVariable.oCompanyInfo.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                End If
                            End If

                            cn.Integration_RunQuery("sp_Invoice_UpdateReceived '" + CStr(HeaderID) + "','" + ret + "'")
                        Else
                            If PublicVariable.oCompanyInfo.InTransaction Then
                                PublicVariable.oCompanyInfo.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            End If
                            If ret.Contains("'") Then
                                ret = ret.Replace(",", "''")
                            End If
                            cn.Integration_RunQuery("sp_Invoice_UpdateReceived '" + CStr(HeaderID) + "','" + ret + "'")
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            If PublicVariable.oCompanyInfo.InTransaction Then
                PublicVariable.oCompanyInfo.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
            Functions.WriteLog(ex.ToString)
        End Try
    End Sub
#End Region
End Class

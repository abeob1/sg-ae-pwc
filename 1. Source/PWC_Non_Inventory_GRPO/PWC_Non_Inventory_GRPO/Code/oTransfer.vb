Public Class oTransfer
#Region "Build Table Structure"
    Private Function BuildTableOWTR() As DataTable
        Dim dt As New DataTable("OWTR")
        dt.Columns.Add("U_POSTxNo")
        dt.Columns.Add("Filler")
        dt.Columns.Add("DocDate")
        dt.Columns.Add("TaxDate")
        dt.Columns.Add("U_ToStore")
        Return dt
    End Function
    Private Function BuildTableWTR1() As DataTable
        Dim dt As New DataTable("WTR1")
        dt.Columns.Add("LineNum")
        dt.Columns.Add("ItemCode")
        dt.Columns.Add("Dscription")
        dt.Columns.Add("WhsCode")
        dt.Columns.Add("Quantity")
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
    Private Function InsertIntoOWTR(dt As DataTable, dr As DataRow) As DataTable
        Dim drNew As DataRow = dt.NewRow
        drNew("U_POSTxNo") = dr("U_POSTxNo")
        drNew("Filler") = dr("Filler")
        drNew("DocDate") = CDate(dr("DocDate")).ToString("yyyyMMdd")
        drNew("TaxDate") = CDate(dr("DocDate")).ToString("yyyyMMdd")
        drNew("U_ToStore") = dr("U_ToStore")
        dt.Rows.Add(drNew)
        Return dt
    End Function
    Private Function InsertIntoWTR1(dt As DataTable, dr As DataRow) As DataTable
        Dim drNew As DataRow = dt.NewRow
        drNew("LineNum") = dr("LineNum")
        drNew("ItemCode") = dr("ItemCode")
        'drNew("Dscription") = dr("Dscription")
        drNew("WhsCode") = dr("WhsCode")
        drNew("Quantity") = dr("Quantity")
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
    Public Sub CreateTransfer()
        Try
            Dim oRunning As New oRunningMonitor
            Dim DocType As String = "67"
            Dim cn As New Connection
            Dim xm As New oXML
            xm.SetDB()
            Dim dt As DataTable = cn.Integration_RunQuery("sp_Transfer_LoadForSync 'Pivotal'")
            If Not IsNothing(dt) Then
                If PublicVariable.oCompanyInfo.Connected = False Then
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
                    oRunning.UpdateMonitor("Transfer", HeaderID)

                    Dim ret As String = ""
                    Dim ds As New DataSet
                    Dim dtOWTR As DataTable = BuildTableOWTR()
                    Dim dtWTR1 As DataTable = BuildTableWTR1()
                    Dim dtSRNT As DataTable = BuildTableSRNT()
                    '----------add header----------
                    dtOWTR = InsertIntoOWTR(dtOWTR, dr)

                    '----------add line------------
                    Dim dtLine As DataTable = cn.Integration_RunQuery("sp_TransferLine_LoadByID " + CStr(HeaderID))
                    If dtLine.Rows.Count = 0 Then
                        ret = "Service Return: Inventory Transfer has no line item."
                        cn.Integration_RunQuery("sp_Transfer_UpdateReceived " + CStr(HeaderID) + ",'" + ret + "'")
                        Continue For
                    Else
                        For Each drLine As DataRow In dtLine.Rows
                            dtWTR1 = InsertIntoWTR1(dtWTR1, drLine)
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
                    cn.Integration_RunQuery("sp_Transfer_UpdateReceived " + CStr(HeaderID) + ",'" + ret + "'")
                Next
            End If
        Catch ex As Exception
            Functions.WriteLog(ex.ToString)
        End Try
    End Sub
End Class

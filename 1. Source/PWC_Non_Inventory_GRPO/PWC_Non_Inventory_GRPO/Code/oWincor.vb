Imports System.IO
Imports System.Data.SqlClient

Public Class oWincor
#Region "Build Datatable"
    Private Function BuildTableHeader() As DataTable
        Dim dt As New DataTable("OPOR")
        dt.Columns.Add("ID", System.Type.GetType("System.Int32"))
        dt.Columns.Add("Receipt_Nmbr")
        dt.Columns.Add("Date")
        dt.Columns.Add("Cust_Name")
        dt.Columns.Add("Phone_Number")
        dt.Columns.Add("Address_1")
        dt.Columns.Add("Address_2")
        dt.Columns.Add("SendDate")
        dt.Columns.Add("ReceiveDate")
        dt.Columns.Add("ErrMsg")
        dt.Columns.Add("Source")
        dt.Columns.Add("UserId", GetType(System.String))
        dt.Columns.Add("SalesPerson", GetType(System.String))
        Return dt
    End Function
    Private Function BuildTableLine() As DataTable
        Dim dt As New DataTable("POR1")
        dt.Columns.Add("HeaderID")
        dt.Columns.Add("Item_Code")
        dt.Columns.Add("Item_Qty")
        dt.Columns.Add("New_Price")
        dt.Columns.Add("Tax_Amount")
        dt.Columns.Add("Item_NSales")
        dt.Columns.Add("OutletCode")
        dt.Columns.Add("IsGST")
        Return dt
    End Function
    Private Function BuildTablePayment() As DataTable
        Dim dt As New DataTable("POR12")
        dt.Columns.Add("HeaderID")
        dt.Columns.Add("Payment_Name")
        dt.Columns.Add("Amount")
        Return dt
    End Function
#End Region
#Region "Insert datatable"
    Private Function InsertIntoTableHeader(dt As DataTable, ID As String, Receipt_Nmbr As String, StrDate As String, _
                                            Cust_Name As String, Phone_Number As String, Address_1 As String, _
                                            Address_2 As String, UserId As String, SalesPerson As String) As DataTable
        Dim drNew As DataRow = dt.NewRow
        drNew("ID") = ID
        drNew("Receipt_Nmbr") = Receipt_Nmbr
        drNew("Date") = StrDate
        drNew("Cust_Name") = Cust_Name
        drNew("Phone_Number") = Phone_Number
        drNew("Address_1") = Address_1
        drNew("Address_2") = Address_2

        drNew("SendDate") = Now
        drNew("ReceiveDate") = DBNull.Value
        drNew("ErrMsg") = DBNull.Value
        drNew("Source") = "Wincor"

        drNew("UserId") = CStr(UserId)
        drNew("SalesPerson") = CStr(SalesPerson)

        dt.Rows.Add(drNew)
        Return dt
    End Function

    Private Function UpdateTableHeader(dt As DataTable, ID As String, Cust_Name As String, Phone_Number As String, _
                                       Address_1 As String, Address_2 As String) As DataTable
        Try
            Dim drNew() As DataRow = dt.Select("ID='" + ID + "'")

            drNew(0)("Cust_Name") = Cust_Name
            drNew(0)("Phone_Number") = Phone_Number
            drNew(0)("Address_1") = Address_1
            drNew(0)("Address_2") = Address_2

            Return dt
        Catch ex As Exception
            MessageBox.Show("UpdateTableHeader: " + ID + "-" + ex.ToString)
            Return dt
        End Try
        
    End Function

    Private Function InsertIntoTableLine(dt As DataTable, HeaderID As String, Item_Code As String, Item_Qty As Integer, _
                                           New_Price As Double, Tax_Amount As Double, Item_NSales As Double, _
                                           OutletCode As String) As DataTable
        Dim drNew As DataRow = dt.NewRow
        drNew("HeaderID") = HeaderID
        drNew("Item_Code") = Item_Code
        drNew("Item_Qty") = Item_Qty
        drNew("New_Price") = New_Price
        drNew("Tax_Amount") = Tax_Amount
        drNew("Item_NSales") = Item_NSales
        drNew("OutletCode") = OutletCode
        drNew("IsGST") = "N"
        dt.Rows.Add(drNew)
        Return dt
    End Function
    Private Function InsertIntoTablePayment(dt As DataTable, HeaderID As String, Payment_Name As String, Amount As Double) As DataTable
        Dim drNew As DataRow = dt.NewRow
        drNew("HeaderID") = HeaderID
        drNew("Payment_Name") = Payment_Name
        drNew("Amount") = Amount
        dt.Rows.Add(drNew)
        Return dt
    End Function
    Private Function GenerateHeaderID(dt As DataTable) As Integer
        If IsNothing(dt) Then Return GetMaxIDfromDB() + 1
        If dt.Rows.Count = 0 Then Return GetMaxIDfromDB() + 1
        Dim MaxID As Integer = 0

        MaxID = dt.Compute("Max(ID)", "")

        Return MaxID + 1
    End Function
    Private Function GetMaxIDfromDB() As Integer
        Dim cn As New Connection
        Dim dt As DataTable
        dt = cn.Integration_RunQuery("Select isnull(Max(ID),0) ID from WC_InvoiceHeader")
        If IsNothing(dt) Then Return 0
        If dt.Rows.Count = 0 Then Return 0
        Return dt.Rows(0).Item("ID")

    End Function
    
#End Region
    Public Function ReadingWincorFile(FilePath As String) As DataSet
        Try
            If File.Exists(FilePath) Then
                Dim ary As Array
                Dim str As String = My.Computer.FileSystem.ReadAllText(FilePath)
                ary = str.Split(vbNewLine)
                Dim dtHeader As DataTable = BuildTableHeader()
                Dim dtLine As DataTable = BuildTableLine()
                Dim dtPayment As DataTable = BuildTablePayment()

                'Header Table
                Dim ID As Integer, Receipt_Nmbr As String, StrDate As String, _
                                            Cust_Name As String, Phone_Number As String, Address_1 As String, _
                                            Address_2 As String, TrainingMode As String = "", UserID As String, _
                                            SalesPerson As String, IsGST As String = "", IsRefund As String = ""
                'Line Table
                Dim Item_Code As String, Item_Qty As Double, _
                                           New_Price As Double, Tax_Amount As Double, Item_NSales As Double, _
                                           OutletCode As String, LineVoid As String = ""
                'Payment Table
                Dim Payment_Name As String, Amount As Double, Type As String

                For Each line In ary 'BY LINE
                    line = line.Replace(vbLf, "")

                    Dim arelement As Array
                    arelement = line.split("|")
                    If arelement.Length > 0 Then
                        Select Case arelement(0).Replace(vbNewLine, "")
                            Case "1"
                                OutletCode = arelement(2)
                            Case "101"
                                Receipt_Nmbr = arelement(1)
                                StrDate = PostingDate(arelement(3), arelement(4))
                                TrainingMode = arelement(12)
                                UserID = arelement(5)
                                SalesPerson = arelement(9)
                                IsRefund = arelement(7)

                                '------------insert invoice header--------------
                                If TrainingMode = "N" Then
                                    ID = GenerateHeaderID(dtHeader)
                                    dtHeader = InsertIntoTableHeader(dtHeader, ID, Receipt_Nmbr, StrDate, "", "", "", "", UserID, SalesPerson)
                                End If
                                
                            Case "103"
                                Cust_Name = arelement(1)
                                Phone_Number = arelement(2)
                                Address_1 = arelement(3)
                                Address_2 = arelement(4)
                                dtHeader = UpdateTableHeader(dtHeader, ID, Cust_Name, Phone_Number, Address_1, Address_2)

                            Case "111"
                                Item_Code = arelement(1)
                                Item_Qty = arelement(2)
                                New_Price = arelement(4)
                                LineVoid = arelement(5)
                                Item_NSales = arelement(13)
                                Tax_Amount = arelement(16)

                                '----------insert invoice line-------------------
                                If LineVoid <> "V" Then
                                    InsertIntoTableLine(dtLine, ID, Item_Code, Item_Qty, New_Price, Tax_Amount, Item_NSales, OutletCode)
                                End If
                            Case "121"
                                IsGST = arelement(7)
                                Try
                                    '-----------update tax ---------------------
                                    If IsGST = "Y" Then 'no tax
                                        For Each dr As DataRow In dtLine.Select("HeaderID='" + CStr(ID) + "'")
                                            dr("IsGST") = IsGST
                                        Next
                                    End If
                                Catch ex As Exception
                                    MessageBox.Show("121-UpdateTax IsGST -" + CStr(ID) + "-" + ex.ToString)
                                End Try
                                

                            Case "131"
                                Payment_Name = arelement(2)
                                Amount = arelement(8)
                                Type = arelement(1)

                                '---------insert payment----------------
                                Dim TypeC As Integer = 1
                                Dim Ref As Integer = 1

                                If Type.ToUpper = "C" Then
                                    TypeC = -1
                                End If
                                If IsRefund <> "" Then
                                    Ref = -1
                                End If

                                Amount = Amount * TypeC * Ref
                                InsertIntoTablePayment(dtPayment, ID, Payment_Name, Amount)
                        End Select
                       
                        'If Item_Code <> "" Then
                        '    If Receipt_Nmbr <> "" And TrainingMode = "N" Then
                        '        ID = GenerateHeaderID(dtHeader)
                        '        dtHeader = InsertIntoTableHeader(dtHeader, ID, Receipt_Nmbr, StrDate, Cust_Name, _
                        '                                        Phone_Number, Address_1, Address_2, UserID, SalesPerson)
                        '        Receipt_Nmbr = ""
                        '        Phone_Number = ""
                        '        Address_1 = ""
                        '        Address_2 = ""
                        '        Cust_Name = ""
                        '        TrainingMode = ""
                        '        UserID = ""
                        '        SalesPerson = ""
                        '        IsRefund = ""
                        '    End If
                        '    If LineVoid <> "V" Then
                        '        InsertIntoTableLine(dtLine, ID, Item_Code, Item_Qty, New_Price, Tax_Amount, Item_NSales, OutletCode)
                        '        Item_Code = ""
                        '        LineVoid = ""
                        '    End If
                        'End If
                        'If Payment_Name <> "" Then
                        '    Dim TypeC As Integer = 1
                        '    Dim Ref As Integer = 1

                        '    If Type.ToUpper = "C" Then
                        '        TypeC = -1
                        '    End If
                        '    If IsRefund <> "" Then
                        '        Ref = -1
                        '    End If

                        '    Amount = Amount * TypeC * Ref
                        '    InsertIntoTablePayment(dtPayment, ID, Payment_Name, Amount)
                        '    Payment_Name = ""
                        '    '-----------update tax ---------------------
                        '    If IsGST = "Y" Then 'no tax
                        '        For Each dr As DataRow In dtLine.Rows
                        '            If dr("HeaderID") = ID Then
                        '                dr("IsGST") = IsGST
                        '            End If
                        '        Next
                        '    End If
                        'End If
                    End If
                Next
                '----remove header without item--------
                dtHeader = DeleteInvoiceWOItem(dtHeader, dtLine)

                '-----update tax logic----------
                dtLine = UpdateTaxPrice(dtLine)

                '-----GROUP PAYMENT TABLE------------
                dtPayment = GroupPaymentTable(dtPayment, dtHeader)

                If dtHeader.Rows.Count > 0 And dtLine.Rows.Count > 0 And dtPayment.Rows.Count > 0 Then
                    Dim ds As New DataSet
                    ds.Tables.Add(dtHeader.Copy)
                    ds.Tables.Add(dtLine.Copy)
                    ds.Tables.Add(dtPayment.Copy)
                    If CheckDuplicate(dtHeader) Then
                        Functions.WriteLog("Duplicate receipt number in same date")
                        MessageBox.Show("Duplicate receipt number in same date")
                        Return Nothing
                    Else
                        Return ds
                    End If

                Else
                    Return Nothing
                End If
            Else
                Return Nothing
            End If

        Catch ex As Exception
            Functions.WriteLog(ex.ToString)
            MessageBox.Show(ex.ToString)
            Return Nothing
        End Try
    End Function
    Public Function BindDStoTable(ds As DataSet) As String
        Try
            If CheckDuplicateInDB(ds.Tables(0)) Then
                Functions.WriteLog("Duplicate receipt number in same date")
                Return "Duplicate receipt number in same date"
            End If
            ds.Tables(1).Columns.Remove("IsGST")
            Dim s As SqlBulkCopy = New SqlBulkCopy(PublicVariable.IntegrationConnectionString)
            s.DestinationTableName = "WC_InvoiceHeader"
            s.WriteToServer(ds.Tables(0))
            s.DestinationTableName = "WC_InvoiceLine"
            s.WriteToServer(ds.Tables(1))
            s.DestinationTableName = "WC_PaymentMean"
            s.WriteToServer(ds.Tables(2))

            Return ""
        Catch ex As Exception
            Functions.WriteLog(ex.ToString)
            Return ex.ToString
        End Try
    End Function
    Private Function CheckDuplicate(dt As DataTable) As Boolean
        Try

       
            Dim dtDate As DataTable
            '1.Get distinct date
            dtDate = dt.DefaultView.ToTable("Date", True, "Date")
            For Each drdate As DataRow In dtDate.Rows
                '2.create table by date
                Dim dtByDate As DataTable = dt.Clone
                For Each dr As DataRow In dt.Select("Date='" + drdate("Date").ToString + "'")
                    dtByDate.ImportRow(dr)
                Next

                '3.Get distinct by Receipt No
                Dim dt1 As DataTable = dtByDate.DefaultView.ToTable(True, "Receipt_Nmbr")
                '4. compare count by date and by receipt no
                If dt1.Rows.Count <> dtByDate.Rows.Count Then
                    Return True
                End If
            Next
            Return False
        Catch ex As Exception
            MessageBox.Show("CheckDuplicate: " + ex.ToString)
            Return False
        End Try
    End Function
    Private Function CheckDuplicateInDB(dt As DataTable) As Boolean
        Dim cn As New Connection
        Dim dt1 As DataTable
        For Each dr As DataRow In dt.Rows
            dt1 = cn.Integration_RunQuery("sp_WCHeader_CheckDuplicate '" + dr("Date").ToString + "','" + dr("Receipt_Nmbr").ToString() + "'")
            If Not IsNothing(dt1) Then
                If dt1.Rows.Count > 0 Then
                    If dt1.Rows(0).Item("cnt") > 0 Then
                        Return True
                    End If
                End If
            End If
        Next

        Return False
    End Function
    Private Function GroupPaymentTable(dtPayment As DataTable, dtHeader As DataTable) As DataTable
        Try
            For Each dr As DataRow In dtHeader.Rows
                Dim newdtpayment As DataRow() = dtPayment.Select(String.Format("HeaderID='{0}' and Payment_Name='CASH'", dr("ID").ToString))
                If newdtpayment.Count > 1 Then
                    Dim drnew As DataRow = dtPayment.NewRow

                    Dim NewAmount As Double = 0
                    For i As Integer = 0 To newdtpayment.Count - 1
                        NewAmount = NewAmount + newdtpayment(i)("Amount")
                        dtPayment.Rows.Remove(newdtpayment(i))
                    Next
                    drnew("HeaderID") = dr("ID").ToString
                    drnew("Payment_Name") = "CASH"
                    drnew("Amount") = NewAmount

                    dtPayment.Rows.Add(drnew)
                End If
            Next

            Return dtPayment
        Catch ex As Exception
            MessageBox.Show("GroupPayment: " + ex.ToString)
            Return dtPayment
        End Try

    End Function

    Private Function UpdateTaxPrice(dtLine As DataTable) As DataTable
        Try

        
            For Each dr As DataRow In dtLine.Rows
                If dr("IsGST") = "Y" Then 'no tax
                    dr("Tax_Amount") = 0
                Else
                    dr("New_Price") = (Double.Parse(dr("Item_NSales")) + Double.Parse(dr("Tax_Amount"))) / Double.Parse(dr("Item_Qty"))
                    dr("Item_NSales") = Double.Parse(dr("Item_NSales")) + Double.Parse(dr("Tax_Amount"))
                End If
            Next
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        Return dtLine
    End Function
    Private Function PostingDate(strdate As String, strtime As String)
        Dim newstrdate As String = strdate
        Dim dFrom As DateTime
        Dim dTo As DateTime
        Dim dCompare As DateTime
        Dim sDateFrom As String = "00:00:00"
        Dim sDateTo As String = "05:00:00"

        If DateTime.TryParse(sDateFrom, dFrom) AndAlso DateTime.TryParse(sDateTo, dTo) AndAlso DateTime.TryParse(strtime, dCompare) Then
            Dim TS As TimeSpan = dTo - dCompare
            If TS.Ticks >= 0 Then
                Dim adate As Date = New Date(CInt(strdate.Substring(0, 4)), CInt(strdate.Substring(4, 2)), CInt(strdate.Substring(6, 2)))
                adate = DateAdd(DateInterval.Day, -1, adate)
                newstrdate = adate.ToString("yyyyMMdd")
            End If
        End If
        Return newstrdate
    End Function

    Private Function DeleteInvoiceWOItem(dtHeader As DataTable, dtLine As DataTable) As DataTable
        Dim newdt As DataTable = dtHeader.Clone
        Try

            For Each dr As DataRow In dtHeader.Rows
                Dim drline() As DataRow = dtLine.Select("HeaderID='" + dr("ID").ToString() + "'")
                If drline.Length > 0 Then
                    newdt.ImportRow(dr)
                End If
            Next
            Return newdt
        Catch ex As Exception
            MessageBox.Show("DeleteInvoiceWOItem :" + ex.ToString)
            Return Newdt
        End Try

    End Function
End Class

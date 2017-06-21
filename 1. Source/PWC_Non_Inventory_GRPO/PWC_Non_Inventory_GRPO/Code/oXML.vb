Imports System.Text
Imports System.Xml

Public Class oXML
    Public Function ToXMLStringFromDS(ObjType As String, ds As DataSet) As String
        Try
            'Dim gf As New GeneralFunctions()
            Dim XmlString As New StringBuilder()
            Dim writer As XmlWriter = XmlWriter.Create(XmlString)
            writer.WriteStartDocument()
            If True Then
                writer.WriteStartElement("BOM")
                If True Then
                    writer.WriteStartElement("BO")
                    If True Then
                        '#Region "write ADMINFO_ELEMENT"
                        writer.WriteStartElement("AdmInfo")
                        If True Then
                            writer.WriteStartElement("Object")
                            If True Then
                                writer.WriteString(ObjType)
                            End If
                            writer.WriteEndElement()
                        End If
                        writer.WriteEndElement()
                        '#End Region

                        '#Region "Header&Line XML"
                        For Each dt As DataTable In ds.Tables
                            If dt.Rows.Count > 0 Then
                                writer.WriteStartElement(dt.TableName.ToString())
                                If True Then
                                    For Each row As DataRow In dt.Rows
                                        writer.WriteStartElement("row")
                                        If True Then
                                            For Each column As DataColumn In dt.Columns
                                                If column.DefaultValue.ToString() <> "xx_remove_xx" Then
                                                    If row(column).ToString() <> "" Then
                                                        writer.WriteStartElement(column.ColumnName)
                                                        'Write Tag
                                                        If True Then
                                                            writer.WriteString(row(column).ToString())
                                                        End If
                                                        writer.WriteEndElement()
                                                    End If
                                                End If
                                            Next
                                        End If
                                        writer.WriteEndElement()
                                    Next
                                End If
                                writer.WriteEndElement()
                            End If
                        Next
                        '#End Region
                    End If
                    writer.WriteEndElement()
                End If
                writer.WriteEndElement()
            End If
            writer.WriteEndDocument()

            writer.Flush()

            Return XmlString.ToString()
        Catch ex As Exception
            Return ex.ToString()
        End Try
    End Function
    Public Function CreateMarketingDocument(ByVal strXml As String, DocType As String) As String
        Try
            Dim sStr As String = ""
            Dim lErrCode As Integer
            Dim sErrMsg As String
            Dim oDocment
            Select Case DocType
                Case "30"
                    oDocment = DirectCast(oDocment, SAPbobsCOM.JournalEntries)
                Case "97"
                    oDocment = DirectCast(oDocment, SAPbobsCOM.SalesOpportunities)
                Case "191"
                    oDocment = DirectCast(oDocment, SAPbobsCOM.ServiceCalls)
                Case "33"
                    oDocment = DirectCast(oDocment, SAPbobsCOM.Contacts)
                Case "221"
                    oDocment = DirectCast(oDocment, SAPbobsCOM.Attachments2)
                Case "2"
                    oDocment = DirectCast(oDocment, SAPbobsCOM.BusinessPartners)
                Case Else
                    oDocment = DirectCast(oDocment, SAPbobsCOM.Documents)
            End Select

            If PublicVariable.oCompanyInfo.Connected = False Then
                SetDB()
                sErrMsg = ConnectSAPDB("")
                If sErrMsg <> "" Then
                    Return sErrMsg
                End If
            End If

            PublicVariable.oCompanyInfo.XMLAsString = True
            oDocment = PublicVariable.oCompanyInfo.GetBusinessObjectFromXML(strXml, 0)

            lErrCode = oDocment.Add()

            If lErrCode <> 0 Then
                PublicVariable.oCompanyInfo.GetLastError(lErrCode, sErrMsg)
                Return sErrMsg
            Else
                Return ""
            End If

        Catch ex As Exception
            Return ex.ToString
        End Try
    End Function
    Public Sub SetDB()
        Try
            Dim MyArr As Array
            Dim Str As String

            '----------Connection to SAP DB------------------
            Str = System.Configuration.ConfigurationSettings.AppSettings.Get("SAPConnectionString")
            MyArr = Str.Split(";")

            PublicVariable.SAPConnectionString = "server= " + MyArr(3).ToString() + ";database=" + MyArr(0).ToString() + " ;uid=" + MyArr(4).ToString() + "; pwd=" + MyArr(5).ToString() + ";"

            If IsNothing(PublicVariable.oCompanyInfo) Then
                PublicVariable.oCompanyInfo = New SAPbobsCOM.Company
            End If
            PublicVariable.oCompanyInfo.CompanyDB = MyArr(0).ToString()
            PublicVariable.oCompanyInfo.UserName = MyArr(1).ToString()
            PublicVariable.oCompanyInfo.Password = MyArr(2).ToString()
            PublicVariable.oCompanyInfo.Server = MyArr(3).ToString()
            PublicVariable.oCompanyInfo.DbUserName = MyArr(4).ToString()
            PublicVariable.oCompanyInfo.DbPassword = MyArr(5).ToString()
            PublicVariable.oCompanyInfo.LicenseServer = MyArr(6)
            Dim SQLType As String = MyArr(7)
            If SQLType = 2008 Then
                PublicVariable.oCompanyInfo.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008
            ElseIf SQLType = 2005 Then
                PublicVariable.oCompanyInfo.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2005
            Else
                PublicVariable.oCompanyInfo.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012
            End If

            '----------------Connection To Integration DB----------------
            Str = System.Configuration.ConfigurationSettings.AppSettings.Get("IntegrationConnectionString")
            MyArr = Str.Split(";")
            PublicVariable.IntegrationConnectionString = "server= " + MyArr(1).ToString() + ";database=" + MyArr(0).ToString() + " ;uid=" + MyArr(2).ToString() + "; pwd=" + MyArr(3).ToString() + ";"

            'Dim oM As New oMapping
            'oM.LoadMapping()
        Catch ex As Exception
            Functions.WriteLog("SetDB:" + ex.ToString)
        End Try
    End Sub
    Public Function ConnectSAPDB(ByVal DBName As String) As String
        Dim lRetCode As Integer
        Dim lErrCode As Integer
        Dim sErrMsg As String = ""
        Try
            If PublicVariable.oCompanyInfo.Connected Then
                PublicVariable.oCompanyInfo.Disconnect()
                'PublicVariable.oCompanyInfo = New SAPbobsCOM.Company
            End If

            Dim MyArr As Array
            Dim Str As String
            'SELECT T0.[U_AB_USERCODE], T0.[U_AB_PASSWORD],T0.[Name] FROM [dbo].[@AB_COMPANYDATA]  T0 WHERE T0.[Name] =''


            Str = "SELECT T0.[U_AB_PASSWORD], T0.[U_AB_USERCODE],T0.[Name] FROM [dbo].[@AB_COMPANYDATA]  T0 WHERE T0.[Name] ='" & DBName & "'"
            Dim cn As New Connection
            Dim dt As DataTable = cn.Integration_RunQuery(Str)
            If dt.Rows.Count = 0 Then
                Return "SAP Suser ID not defined in AB_COMPANYDATA Data"
            End If
            Dim SAPPwd As String = dt.Rows(0).Item("U_AB_PASSWORD").ToString
            Dim SAPUser As String = dt.Rows(0).Item("U_AB_USERCODE").ToString

            '----------Connection to SAP DB------------------
            Str = System.Configuration.ConfigurationSettings.AppSettings.Get("SAPConnectionString")
            MyArr = Str.Split(";")

            If IsNothing(PublicVariable.oCompanyInfo) Then
                PublicVariable.oCompanyInfo = New SAPbobsCOM.Company
            End If
            PublicVariable.oCompanyInfo.CompanyDB = DBName 'MyArr(0).ToString()
            PublicVariable.oCompanyInfo.UserName = SAPUser 'MyArr(1).ToString()
            PublicVariable.oCompanyInfo.Password = SAPPwd 'MyArr(2).ToString()
            PublicVariable.oCompanyInfo.Server = MyArr(3).ToString()
            PublicVariable.oCompanyInfo.DbUserName = MyArr(4).ToString()
            PublicVariable.oCompanyInfo.DbPassword = MyArr(5).ToString()
            PublicVariable.oCompanyInfo.LicenseServer = MyArr(6)
            Dim SQLType As String = MyArr(7)
            'If SQLType = 2008 Then
            If SQLType = "2008" Then
                PublicVariable.oCompanyInfo.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008
            ElseIf SQLType = "2005" Then
                PublicVariable.oCompanyInfo.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2005
            Else
                PublicVariable.oCompanyInfo.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012
            End If

            lRetCode = PublicVariable.oCompanyInfo.Connect
            If lRetCode <> 0 Then
                PublicVariable.oCompanyInfo.GetLastError(lErrCode, sErrMsg)
                Functions.WriteLog(sErrMsg & "DB:" & DBName)
                Return sErrMsg
            Else
                Functions.WriteLog("ConnectSAPDB OK" & "DB:" & DBName)
                Return ""
            End If
        Catch ex As Exception
            Functions.WriteLog("ConnectSAPDB " + sErrMsg & "DB:" & DBName)
            Return sErrMsg
        End Try
    End Function
End Class

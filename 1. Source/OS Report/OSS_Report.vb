Imports System.Diagnostics.Process
Imports System.Threading
Imports System.IO
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Public Class OSS_Report
    Dim WithEvents SBO_Application As SAPbouiCOM.Application
    Dim Ocompany As SAPbobsCOM.Company
    Sub New(ByVal ocompany1 As SAPbobsCOM.Company, ByVal sbo_application1 As SAPbouiCOM.Application)
        SBO_Application = sbo_application1
        Ocompany = ocompany1
    End Sub

    Public Sub Form_Bind(ByVal oForm As SAPbouiCOM.Form)
        Try
          
            oForm.DataSources.UserDataSources.Add("oedit1", SAPbouiCOM.BoDataType.dt_DATE)
            oEdit = oForm.Items.Item("3").Specific
            oEdit.DataBind.SetBound(True, "", "oedit1")

        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try


    End Sub


    Private Sub SBO_Application_ItemEvent(FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        Try
            If pVal.FormUID = "OSS" Then
                Try
                    oForm = SBO_Application.Forms.Item("OSS")
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.BeforeAction = False And pVal.InnerEvent = False Then
                        If pVal.ItemUID = "Ex" Then
                            Dim Fdate As String = ""
                            oEdit = oForm.Items.Item("3").Specific
                            Fdate = oEdit.String
                            If Fdate = "" Then
                                SBO_Application.StatusBar.SetText("Date Can't Be Empty!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Exit Sub
                            End If
                            Dim trd As Threading.Thread
                            trd = New Threading.Thread(AddressOf ConslInvReport1)
                            trd.IsBackground = True
                            trd.SetApartmentState(ApartmentState.STA)
                            trd.Start()

                        End If
                    End If

                Catch ex As Exception

                End Try

            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ConslInvReport1()
        Try
            oForm = SBO_Application.Forms.Item("OSS")
            Dim ConDate As String
          
            oEdit = oForm.Items.Item("3").Specific
            ConDate = oEdit.String 'Format(oEdit.Value, "yyyy-MM-dd")
            If ConDate = "" Then
                SBO_Application.StatusBar.SetText("Date Can't Be Empty!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            'dd/mm/yyyy
            'ConDate = ConDate.Substring(0, 2)
            'ConDate = ConDate.Substring(3, 2)
            ConDate = ConDate.Substring(6, 4) + "-" + ConDate.Substring(3, 2) + "-" + ConDate.Substring(0, 2)
            Dim DocEntry As String = ""
            Dim i As Integer = 0

            Dim sPath As String
            sPath = IO.Directory.GetParent(Application.StartupPath).ToString
            Dim file As System.IO.StreamReader = New System.IO.StreamReader(sPath & "\PWC\" & "Pwd.txt", True)
            Dim pwd As String
            pwd = file.ReadLine()
            Dim file1 As System.IO.StreamReader = New System.IO.StreamReader(sPath & "\PWC\" & "UID.txt", True)
            Dim UID As String
            UID = file1.ReadLine()
            Dim file2 As System.IO.StreamReader = New System.IO.StreamReader(sPath & "\PWC\" & "Path.txt", True)
            Dim path As String
            path = file2.ReadLine()
            Dim file3 As System.IO.StreamReader = New System.IO.StreamReader(sPath & "\PWC\" & "Serv.txt", True)
            Dim Servername As String
            Servername = file3.ReadLine()
            SBO_Application.StatusBar.SetText(ConDate, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Dim Los As String = "15ABCD"
            Dim Los1 As String = "15ABCD"
            Dim oRecordSet As SAPbobsCOM.Recordset
            Dim sqlstr As String = "SELECT T0.[PrcCode],T0.[PrcName] FROM OPRC T0 WHERE T0.[DimCode] =3 and  T0.[Active] ='Y'"
            oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(sqlstr)
            If oRecordSet.RecordCount > 0 Then
                For i = 0 To oRecordSet.RecordCount - 1

                    Los = oRecordSet.Fields.Item(0).Value.ToString
                    Los1 = oRecordSet.Fields.Item(1).Value.ToString
                    Dim cryRpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument
                    Dim crtableLogoninfos As New CrystalDecisions.Shared.TableLogOnInfos
                    Dim crtableLogoninfo As New CrystalDecisions.Shared.TableLogOnInfo
                    Dim crConnectionInfo As New CrystalDecisions.Shared.ConnectionInfo
                    Dim CrTables As CrystalDecisions.CrystalReports.Engine.Tables
                    Dim CrTable As CrystalDecisions.CrystalReports.Engine.Table
                   
                    cryRpt.Load(sPath & "\PWC\NEW_OR_OU.rpt")
                    'Dim crParameterFieldDefinitions As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinitions
                    'Dim crParameterFieldDefinition As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinition
                    Dim crParameterValues As New CrystalDecisions.Shared.ParameterValues
                    Dim crParameterDiscreteValue As New CrystalDecisions.Shared.ParameterDiscreteValue

                    cryRpt.SetParameterValue(0, ConDate)
                    cryRpt.SetParameterValue(1, ConDate)
                    cryRpt.SetParameterValue(2, Los)
                    cryRpt.SetParameterValue(3, Los1)

                    Dim fileName As String = ""
                    Dim Server As String = Ocompany.Server
                    Dim DB As String = Ocompany.CompanyDB

                    fileName = Los1 + "_" + Los + "_OU_" + Format(DateTime.Now, "yyyyMM") + ".xls" '_hhmmss_ffff

                    With crConnectionInfo
                        .ServerName = Servername
                        .DatabaseName = DB
                        .UserID = UID
                        .Password = pwd
                    End With

                    CrTables = cryRpt.Database.Tables
                    For Each CrTable In CrTables
                        crtableLogoninfo = CrTable.LogOnInfo
                        crtableLogoninfo.ConnectionInfo = crConnectionInfo
                        CrTable.ApplyLogOnInfo(crtableLogoninfo)
                    Next

                    'Try
                    Dim CrExportOptions As ExportOptions
                    Dim CrDiskFileDestinationOptions As New  _
                    DiskFileDestinationOptions()
                    Dim CrFormatTypeOptions As New ExcelFormatOptions
                    CrDiskFileDestinationOptions.DiskFileName = path + fileName

                    ' "c:\crystalExport.xls"

                    CrExportOptions = cryRpt.ExportOptions
                    With CrExportOptions
                        .ExportDestinationType = ExportDestinationType.DiskFile
                        .ExportFormatType = ExportFormatType.Excel
                        .DestinationOptions = CrDiskFileDestinationOptions
                        .FormatOptions = CrFormatTypeOptions
                    End With
                    cryRpt.Export()
                    SBO_Application.StatusBar.SetText("Export Started:OU-" & Los1 & " Line No-" & i + 1 & " of " & oRecordSet.RecordCount & " ", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    cryRpt.Dispose()
                    cryRpt = Nothing
                    GC.Collect()
                    oRecordSet.MoveNext()
                Next
            End If

            SBO_Application.StatusBar.SetText("Export Complected", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            'Catch ex As Exception
            '    SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            'End Try

            'Dim RptFrm As MY_Report
            'RptFrm = New MY_Report
            'RptFrm.CrystalReportViewer1.ReportSource = cryRpt
            'RptFrm.CrystalReportViewer1.Refresh()
            'RptFrm.Text = "OSS Report"
            'RptFrm.TopMost = True

            'RptFrm.Activate()
            'RptFrm.ShowDialog()


        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    ''    crParameterDiscreteValue.Value = (ConDate)
    ''    crParameterFieldDefinitions = _
    ''cryRpt.DataDefinition.ParameterFields
    ''    crParameterFieldDefinition = _
    ''crParameterFieldDefinitions.Item("@FromDt")
    ''    crParameterValues = crParameterFieldDefinition.CurrentValues
    ''    crParameterValues.Clear()
    ''    crParameterValues.Add(crParameterDiscreteValue)
    ''    '    crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

    ''    crParameterDiscreteValue.Value = (ConDate)
    ''    crParameterFieldDefinitions = _
    ''cryRpt.DataDefinition.ParameterFields
    ''    crParameterFieldDefinition = _
    ''crParameterFieldDefinitions.Item("@ToDate")
    ''    crParameterValues = crParameterFieldDefinition.CurrentValues
    ''    crParameterValues.Clear()
    ''    crParameterValues.Add(crParameterDiscreteValue)
    ''    '    crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

    ''    crParameterDiscreteValue.Value = ("")
    ''    crParameterFieldDefinitions = _
    ''cryRpt.DataDefinition.ParameterFields
    ''    crParameterFieldDefinition = _
    ''crParameterFieldDefinitions.Item("@LOS")
    ''    crParameterValues = crParameterFieldDefinition.CurrentValues

    ''    crParameterValues.Clear()
    ''    crParameterValues.Add(crParameterDiscreteValue)
    ''    crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)
End Class

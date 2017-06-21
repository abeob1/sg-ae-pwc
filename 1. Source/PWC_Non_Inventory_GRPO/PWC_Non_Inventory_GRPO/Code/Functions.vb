Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.ServiceProcess
Imports Microsoft.Reporting.WinForms
Imports System.Net.Mail
Imports System.IO.Packaging
Imports System
Imports System.Text
Imports System.Web.Services.Protocols
Imports System.Xml
Imports System.Xml.Serialization
Imports System.Net

Public Class Functions
    
    
    Public Shared Sub WriteLog(ByVal Str As String)
        Dim oWrite As IO.StreamWriter
        Dim FilePath As String
        FilePath = Application.StartupPath + "\logfile.txt"

        If IO.File.Exists(FilePath) Then
            oWrite = IO.File.AppendText(FilePath)
        Else
            oWrite = IO.File.CreateText(FilePath)
        End If
        oWrite.Write(Now.ToString() + ":" + Str + vbCrLf)
        oWrite.Close()
    End Sub
    
    Public Shared Sub WriteXMLLog(DocType As String, XMLStr As String, ErrMsg As String)
        Dim cn As New Connection
        If XMLStr.Contains("'") Then
            XMLStr = XMLStr.Replace(",", "''")
        End If
        If ErrMsg.Contains("'") Then
            ErrMsg = ErrMsg.Replace(",", "''")
        End If
        cn.Integration_RunQuery("sp_XMLLog_Insert '" + DocType + "','" + XMLStr + "','" + ErrMsg + "'")
    End Sub
    Public Sub AutoRun()
        Dim xm As New oXML
        xm.SetDB()

        Try
            PublicVariable.AutoRetry = CBool(System.Configuration.ConfigurationSettings.AppSettings.Get("AutoRetry"))
        Catch ex As Exception
            WriteLog("AutoRun+set autoretry:" + ex.ToString)
        End Try



        'If PublicVariable.AutoRetry Then
        '    Dim a As New oAutoRetry
        '    a.RetryAll()
        'End If

        ''------------RECEIVE GRPO--------------
        'Dim orp As New oGRPO
        'orp.CreateGRPO()

        ''------------RECEIVE TRANFER--------------
        'Dim otr As New oTransfer
        'otr.CreateTransfer()

        ''------------RECEIVE INVOICE--------------
        'Dim oin As New oInvoice
        'oin.CreateInvoice()

        ''------------RECEIVE Goods Return--------------
        'Dim ort As New oReturn
        'ort.CreateGoodsReturn()

        ''------------SEND EMAIL-------------------
        'Dim oem As New oEmailWeb
        'oem.SendPOEmail()



    End Sub
End Class

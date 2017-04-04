Option Strict Off
Option Explicit On
Friend Class HelloWorld
    '//  SAP MANAGE UIDI API 2007 SDK Sample
    '//****************************************************************************
    '//
    '//  File:      HelloWorld.cls
    '//
    '//  Copyright (c) SAP MANAGE
    '//
    '// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
    '// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
    '// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
    '// PARTICULAR PURPOSE.
    '//
    '//****************************************************************************



    '//**************************************************************************************************
    '// BEFORE STARTING:
    '// 1. Add reference to the "SAP Business One UI API"
    '// 2. Add reference to the "SAP Business One Objects Bridge VERSION 2007"
    '// 3. Insert the development connection string to the "Command Line Argument"
    '//-----------------------------------------------------------------
    '// 1.
    '//    a. Project->References
    '//    b. check the "SAP Business One UI API 2007" check box
    '//
    '// 2.
    '//    a. Project->References
    '//    b. check the "SAP Business One Objects Bridge VERSION 2007" check box
    '//
    '// 3.
    '//     a. Project->Properties
    '//     b. choose 'Make' tab folder
    '//     c. place the following connection string in the 'Command Line Arguments' field
    '// 0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056
    '//
    '//**************************************************************************************************


    '//**********************************************************
    '// This parameter will use us to manipulate the
    '// SAP Business One Application a
    '// and the Company on which we are working
    '//**********************************************************

    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private oCompany As SAPbobsCOM.Company

    Private Sub SetApplication()

        '*******************************************************************
        '// Use an SboGuiApi object to establish connection
        '// with the SAP Business One application and return an
        '// initialized appliction object
        '*******************************************************************

        Dim SboGuiApi As SAPbouiCOM.SboGuiApi
        Dim sConnectionString As String

        SboGuiApi = New SAPbouiCOM.SboGuiApi

        '// by following the steps specified above, the following
        '// statment should be suficient for either development or run mode

        sConnectionString = Environment.GetCommandLineArgs.GetValue(1)

        '// connect to a running SBO Application

        SboGuiApi.Connect(sConnectionString)

        '// get an initialized application object

        SBO_Application = SboGuiApi.GetApplication()

    End Sub

    Private Function SetConnectionContext() As Integer

        Dim sCookie As String
        Dim sConnectionContext As String
        Dim lRetCode As Integer

        '// First initialize the Company object

        oCompany = New SAPbobsCOM.Company

        '// Acquire the connection context cookie from the DI API.
        sCookie = oCompany.GetContextCookie

        '// Retrieve the connection context string from the UI API using the
        '// acquired cookie.
        sConnectionContext = SBO_Application.Company.GetConnectionContext(sCookie)

        '// before setting the SBO Login Context make sure the company is not
        '// connected

        If oCompany.Connected = True Then
            oCompany.Disconnect()
        End If

        '// Set the connection context information to the DI API.
        SetConnectionContext = oCompany.SetSboLoginContext(sConnectionContext)

    End Function

    Private Function ConnectToCompany() As Integer

        '// Establish the connection to the company database.
        ConnectToCompany = oCompany.Connect

    End Function
    'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'
    Private Sub Class_Initialize_Renamed()

        '//*************************************************************
        '// set SBO_Application with an initialized application object
        '//*************************************************************

        SetApplication()

        '//*************************************************************
        '// Set The Connection Context
        '//*************************************************************

        If Not SetConnectionContext() = 0 Then
            SBO_Application.MessageBox("Failed setting a connection to DI API")
            End ' Terminating the Add-On Application
        End If


        '//*************************************************************
        '// Connect To The Company Data Base
        '//*************************************************************

        If Not ConnectToCompany() = 0 Then
            SBO_Application.MessageBox("Failed connecting to the company's Data Base")
            End ' Terminating the Add-On Application
        End If

        '//*************************************************************
        '// send an "hello world" message
        '//*************************************************************

        SBO_Application.MessageBox("DI Connected To: " & oCompany.CompanyName & vbNewLine & "Hello World!")

    End Sub
    Public Sub New()
        MyBase.New()
        Class_Initialize_Renamed()
    End Sub
End Class
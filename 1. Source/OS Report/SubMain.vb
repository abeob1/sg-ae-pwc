Option Strict Off
Option Explicit On
Module SubMain
    '//  SAP MANAGE UIDI API 2007 SDK Sample
	'//****************************************************************************
	'//
	'//  File:      SubMain.bas
	'//
	'//  Copyright (c) SAP MANAGE
	'//
	'// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
	'// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
	'// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
	'// PARTICULAR PURPOSE.
	'//
	'//****************************************************************************
	
	
	Public Sub Main()
        Try
            Dim oConnection As HelloWorld
            oConnection = New HelloWorld
            System.Windows.Forms.Application.Run()
        Catch ex As Exception
        End Try

        'Dim oHelloWorld As HelloWorld

        'oHelloWorld = New HelloWorld
		
	End Sub
End Module
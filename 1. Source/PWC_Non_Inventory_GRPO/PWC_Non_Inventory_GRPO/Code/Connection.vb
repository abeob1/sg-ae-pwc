Imports System.Data.SqlClient

Public Class Connection
#Region "ADO Integration"
    Private Function Integration_OpenSQLConnection() As Boolean
        PublicVariable.IntegrationConnection = New SqlConnection
        PublicVariable.IntegrationConnection.ConnectionString = PublicVariable.IntegrationConnectionString

        If PublicVariable.IntegrationConnection.State = ConnectionState.Open Then
            PublicVariable.IntegrationConnection.Close()
        End If
        Try
            PublicVariable.IntegrationConnection.Open()
        Catch ex As Exception
            Functions.WriteLog("Integration_OpenSQLConnection:" + ex.Message)
            Return False
        End Try
        Return True
    End Function
    Public Function Integration_RunQuery(ByVal querystr As String) As DataTable
        If querystr = "" Then Return Nothing

        Try
            Dim MyArr As Array
            Dim Str As String
            Str = System.Configuration.ConfigurationSettings.AppSettings.Get("IntegrationConnectionString")
            MyArr = Str.Split(";")

            Dim NConnection = New SqlConnection
            NConnection.ConnectionString = "server= " + MyArr(1).ToString() + ";database=" + MyArr(0).ToString() + " ;uid=" + MyArr(2).ToString() + "; pwd=" + MyArr(3).ToString() + ";"
            NConnection.Open()

            Dim MyCommand As SqlCommand = New SqlCommand(querystr, NConnection)
            MyCommand.CommandType = CommandType.Text
            MyCommand.CommandTimeout = 0

            Dim da As SqlDataAdapter = New SqlDataAdapter()
            Dim mytable As DataTable = New DataTable()
            da.SelectCommand = MyCommand
            da.Fill(mytable)
            NConnection.Close()
            If mytable Is Nothing Then Return Nothing
            Return mytable
        Catch ex As Exception
            Functions.WriteLog("Integration_RunQuery: " + querystr + " ERROR:" + ex.Message)
            Return Nothing
        End Try
    End Function
    Public Function Integration_RunQuery_BR(ByVal querystr As String, ByVal DBName As String) As DataTable
        If querystr = "" Then Return Nothing

        Try
            Dim MyArr As Array
            Dim Str As String
            Str = System.Configuration.ConfigurationSettings.AppSettings.Get("IntegrationConnectionString")
            MyArr = Str.Split(";")

            Dim NConnection = New SqlConnection
            NConnection.ConnectionString = "server= " + MyArr(1).ToString() + ";database=" + DBName.ToString() + " ;uid=" + MyArr(2).ToString() + "; pwd=" + MyArr(3).ToString() + ";"
            NConnection.Open()

            Dim MyCommand As SqlCommand = New SqlCommand(querystr, NConnection)
            MyCommand.CommandType = CommandType.Text
            MyCommand.CommandTimeout = 0

            Dim da As SqlDataAdapter = New SqlDataAdapter()
            Dim mytable As DataTable = New DataTable()
            da.SelectCommand = MyCommand
            da.Fill(mytable)
            NConnection.Close()
            If mytable Is Nothing Then Return Nothing
            Return mytable
        Catch ex As Exception
            Functions.WriteLog("Integration_RunQuery: " + querystr + " ERROR:" + ex.Message)
            Return Nothing
        End Try
    End Function
#End Region
#Region "ADO SAP"
    Private Function SAP_OpenSQLConnection() As Boolean
        PublicVariable.SAPConnection = New SqlConnection
        Dim MyArr As Array
        Dim Str As String
        Str = System.Configuration.ConfigurationSettings.AppSettings.Get("SAPConnectionString")
        MyArr = Str.Split(";")
        Dim constr As String = "server= " + MyArr(3).ToString() + ";database=" + MyArr(0).ToString() + " ;uid=" + MyArr(4).ToString() + "; pwd=" + MyArr(5).ToString() + ";"

        PublicVariable.SAPConnection.ConnectionString = constr 'PublicVariable.SAPConnectionString

        If PublicVariable.SAPConnection.State = ConnectionState.Open Then
            PublicVariable.SAPConnection.Close()
        End If
        Try
            PublicVariable.SAPConnection.Open()
        Catch ex As Exception
            Functions.WriteLog("SAP_OpenSQLConnection:" + ex.Message)
            Return False
        End Try
        Return True
    End Function
    Public Function SAP_RunQuery(ByVal querystr As String) As DataTable
        Try
            If SAP_OpenSQLConnection() Then
                Dim MyCommand As SqlCommand = New SqlCommand(querystr, PublicVariable.SAPConnection)
                MyCommand.CommandType = CommandType.Text
                MyCommand.CommandText = querystr
                Dim da As SqlDataAdapter = New SqlDataAdapter()
                Dim mytable As DataTable = New DataTable()
                da.SelectCommand = MyCommand
                da.Fill(mytable)
                PublicVariable.SAPConnection.Close()
                If mytable Is Nothing Then Return Nothing
                Return mytable
            Else
                Return Nothing
            End If
        Catch ex As Exception
            Functions.WriteLog("SAP_RunQuery:" + ex.Message)
            Return Nothing
        End Try
    End Function
#End Region
End Class


Imports System.IO

Namespace AE_PWC_AO02

    Module modOSSFileGeneration

        Public Function Write_TextFile(ByVal oDT_FinalResult As DataTable, ByVal sPAth As String, ByVal scheck As String, ByRef sErrDesc As String) As Long
            Try
                Dim sFuncName As String = String.Empty
                Dim irow As Integer
                Dim sFileName As String = "\OSSFile.txt"
                Dim sbuffer As String = String.Empty

                sFuncName = "Write_TextFile()"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

                If File.Exists(sPAth & sFileName) Then
                    Try
                        File.Delete(sPAth & sFileName)
                    Catch ex As Exception
                    End Try
                End If

                Dim sw As StreamWriter = New StreamWriter(sPAth & sFileName)
                ' Add some text to the file.

                For imjs = 0 To oDT_FinalResult.Rows.Count - 1

                    Dim hh As String = "" & oDT_FinalResult.Rows(imjs).Item("AcctCode").ToString & ""

                    sw.WriteLine(oDT_FinalResult.Rows(imjs).Item("AcctCode").ToString & "," & oDT_FinalResult.Rows(imjs).Item("RefDate").ToString & "," & oDT_FinalResult.Rows(imjs).Item("OU Code").ToString & "," _
                                 & oDT_FinalResult.Rows(imjs).Item("Entity").ToString & "," & oDT_FinalResult.Rows(imjs).Item("DC").ToString & "," & oDT_FinalResult.Rows(imjs).Item("Amount").ToString & ",")

                Next imjs
                sw.Close()
                If scheck = "Y" Then
                    Process.Start(sPAth & sFileName)
                End If
                '' Process.Start(sPath & sFileName)

                Write_TextFile = RTN_SUCCESS
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With SUCCESS ", sFuncName)

            Catch ex As Exception
                Write_TextFile = RTN_ERROR
                sErrDesc = ex.Message
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            End Try

        End Function

    End Module

End Namespace



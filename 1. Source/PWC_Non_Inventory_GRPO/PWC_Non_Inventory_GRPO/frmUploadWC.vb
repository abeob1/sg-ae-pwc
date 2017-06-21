Public Class frmUploadWC
    Dim ds As DataSet
    Private Sub btnBrowseFile_Click(sender As System.Object, e As System.EventArgs) Handles btnBrowseFile.Click
        Try
            Dim filename As String
            OpenFileDialog1.Title = "Select Wincor Text File"
            OpenFileDialog1.InitialDirectory = "C:\"
            OpenFileDialog1.Filter = "Text File | *.txt"

            If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                filename = OpenFileDialog1.FileName
                txtFileName.Text = filename
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Me.Cursor = Cursors.WaitCursor
        'Dim owc As New oWincor
        'ds = owc.ReadingWincorFile(txtFileName.Text)
        'If Not IsNothing(ds) Then
        '    grData.DataSource = ds.Tables(0)
        'End If

        Me.Cursor = Cursors.Default
    End Sub

    Private Sub btnGenerate_Click(sender As System.Object, e As System.EventArgs) Handles btnGenerate.Click
        Me.Cursor = Cursors.WaitCursor
        'If Not IsNothing(ds) Then
        '    Dim owc As New oWincor
        '    Dim str As String = ""
        '    str = owc.BindDStoTable(ds)

        '    If str = "" Then
        '        MessageBox.Show("Upload completed!")
        '    Else
        '        MessageBox.Show(str)
        '    End If

        '    ds = Nothing
        '    grData.DataSource = Nothing
        'End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub grData_MouseDoubleClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles grData.MouseDoubleClick
        Try
            If grData.RowCount = 0 Then
                Return
            End If
            If grData.SelectedRows.Count = 0 Then Return

            Dim HeaderID As String = ""
            HeaderID = grData.SelectedRows.Item(0).Cells("ID").Value

            Dim dt As DataTable = ds.Tables(2).Copy
            Dim dt1 As DataTable = dt.Clone
            If HeaderID <> "" Then
                Dim strQuery As String = ""
                strQuery = "HeaderID='" + HeaderID + "'"
                For Each dr As DataRow In dt.Select(strQuery)
                    dt1.ImportRow(dr)
                Next

                Dim frm As New frmPayment
                frm.grMonitor.DataSource = dt1
                frm.ShowDialog()
            End If


            
        Catch ex As Exception

        End Try
    End Sub

    Private Sub grData_SelectionChanged(sender As System.Object, e As System.EventArgs) Handles grData.SelectionChanged
        If grData.SelectedRows.Count = 0 Then Return

        Dim HeaderID As String = ""
        HeaderID = grData.SelectedRows.Item(0).Cells("ID").Value

        Dim dt As DataTable = ds.Tables(1).Copy
        Dim dt1 As DataTable = dt.Clone
        If HeaderID <> "" Then
            Dim strQuery As String = ""
            strQuery = "HeaderID='" + HeaderID + "'"
            For Each dr As DataRow In dt.Select(strQuery)
                dt1.ImportRow(dr)
            Next
            DataGridView1.DataSource = dt1
        End If
    End Sub

    Private Sub grData_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles grData.CellContentClick

    End Sub
End Class
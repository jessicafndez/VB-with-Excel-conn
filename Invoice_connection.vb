Public Class Invoice

    Private Sub Button1_Click(ByVal sender As System.Object, _
                ByVal e As System.EventArgs) Handles Button1.Click
        Try
            Dim MyConnection As System.Data.OleDb.OleDbConnection
            Dim DtSet As System.Data.DataSet
            Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
            MyConnection = New System.Data.OleDb.OleDbConnection _
             ("provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + _
            "'C:\******.xls';Extended Properties=Excel 8.0;")
            MyCommand = New System.Data.OleDb.OleDbDataAdapter _
                ("select * from [Sheet1$]", MyConnection)
            MyCommand.TableMappings.Add("Table", "TestTable")
            DtSet = New System.Data.DataSet
            MyCommand.Fill(DtSet)
            DataGridView1.DataSource = DtSet.Tables(0)
            MyConnection.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

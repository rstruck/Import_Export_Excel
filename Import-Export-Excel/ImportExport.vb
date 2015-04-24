Imports System.Data
Imports System.Data.SqlClient
Imports System.Text
Public Class ImportExport
    Public conn As New SqlConnection("Data Source=SS-SQL\SQL2012;Initial Catalog=master;Integrated Security=True")
    Private Property Command As SqlCommand
    Public MyCommand As System.Data.OleDb.OleDbDataAdapter
    Private Property MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; " & _
                      "data source='"" '; " & "Extended Properties=""Excel 8.0;HDR=NO;IMEX=1;""")
#Region "GeneralUtilities"

    Public Sub ExcelConnections()

    End Sub

    Public Function ADOConnections(ByRef workingdatabase As String) As SqlConnection
        If workingdatabase = Nothing Then
            workingdatabase = "SSDEV"
        End If

        Try
            Dim connectString As String = _
             "Data Source=SS-SQL\SQL2012;" & _
             "Integrated Security=True"

            Dim builder As New SqlConnectionStringBuilder(connectString)
            builder("Database") = workingdatabase

            Dim Newconnectionstring As New SqlConnection(builder.ConnectionString)
            ADOConnections = Newconnectionstring

        Catch ex As Exception
            MessageBox.Show("the database is wrong!!!")
            ADOConnections = conn
        End Try


    End Function


    'Public Function GetFileName() As String


    'End Function

    Public Sub commandobject()

    End Sub

#End Region
End Class

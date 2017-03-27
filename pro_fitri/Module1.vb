Imports System.Data.OleDb
Module Module1
    Public con As OleDbConnection
    Public ds As DataSet
    Public cmd As OleDbCommand
    Public da As OleDbDataAdapter
    Public rd As OleDbDataReader
    Public sql As String

    Sub koneksi()
        Try
            con = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\pro_fitri.accdb")
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
        Catch ex As Exception
            MsgBox("Koneksi Ke database bermasalah, perikasa modul koneksi")
        End Try
    End Sub

End Module

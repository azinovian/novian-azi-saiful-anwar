Public Class Form1
    '---Sub tampilekstra(ByVal sql As String, ByVal nm_table As String, ByVal a As DataGridView)
    '--    Call koneksi()
    '-- da = New OleDb.OleDbDataAdapter(sql, con)
    '--ds = New DataSet
    '--ds.Clear()
    '---da.Fill(ds, nm_table)
    '--a.DataSource = (ds.Tables(nm_table))

    '--End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Form2.Show()
        'Call koneksi()
        'da = New OleDb.OleDbDataAdapter("select * from pasien", con)
        'ds = New DataSet
        'ds.Clear()
        'da.Fill(ds, "pasien")
        'Form2.DataGridView1.DataSource = (ds.Tables("pasien"))

        Me.Hide()
        MsgBox("masukkan terlebih dahulu kode pasien")


    End Sub
End Class

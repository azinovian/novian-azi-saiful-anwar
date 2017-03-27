Imports System.Data.OleDb

Public Class Form2
    Dim DATABARU As Boolean
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click

        Form1.Show()
        Me.Close()

    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call koneksi() 'Memanggil koneksi yang disimpan di module
        '--Proses pengeccekan data pasien--
        'Pengecekkan dengan pemanggilan WHERE---
        Try
            Call koneksi()
            Dim str As String
            str = "select * from pasien where kode_pasien='" & TextBox1.Text & "'"
            cmd = New OleDb.OleDbCommand(str, con)
            rd = cmd.ExecuteReader
            rd.Read()
            If rd.HasRows Then
                Label2.Enabled = False
                TextBox1.Enabled = False
                Button1.Enabled = False

                TextBox2.Text = rd.Item(1)
                TextBox3.Text = rd.Item(2)
                TextBox4.Text = rd.Item(3)
            End If
        Catch ex As Exception

        End Try
        '---Tamat 1.----

        'Proses menampilkan data tindakan
        ComboBox1.Enabled = True
        Try
            Dim str2 As String
            Dim s As String
            str2 = "select kode_tindakan, nama_tindakan from tbl_tindakan"
            cmd = New OleDb.OleDbCommand(str2, con)
            rd = cmd.ExecuteReader

            ComboBox1.Text = "Pilih Tindakan"

            While rd.Read()

                ComboBox1.Items.Add(rd("kode_tindakan"))


                s = rd.Item(0)


            End While
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Me.Close()
        Form1.Close()

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.Text <> "" And ComboBox1.Text <> "Pilih Tindakan" Then

            Call koneksi()
            Dim str5 As String
            str5 = "select * from tbl_tindakan where kode_tindakan='" & ComboBox1.Text & "'"
            cmd = New OleDb.OleDbCommand(str5, con)
            rd = cmd.ExecuteReader
            rd.Read()
            If rd.HasRows Then
                TextBox6.Text = rd.Item(1)
                TextBox5.Text = "Rp. " & rd.Item(2)
            End If

            Button2.Enabled = True

        End If

    End Sub
    'FUNGSI UNTUK SIMPAN
    Sub tblsimpan(ByVal isian As Boolean, ByVal query As String, ByVal query3 As String, ByVal nm_table As String, ByVal datgried As DataGridView)
        Dim simpan As String
        Dim pesan As Integer
        DATABARU = True
        simpan = query
        If isian Then Exit Sub

        If DATABARU Then
            pesan = MsgBox("Apakah anda yakin data akan ditambahkan ke database ?", vbYesNo + vbInformation, "Perhatian")
            If pesan = vbNo Then
                Exit Sub
            End If
            simpan = query
        End If

        Me.Cursor = Cursors.WaitCursor
        jalankansql(simpan) 'memanggil fungsi jalankansql
        datgried.Refresh() 'meng refresh datagried
        'isigrid(query3, nm_table, datgried)
        tampil5()

        Me.Cursor = Cursors.Default
    End Sub
    'Menjalankan fungsi sql
    Private Sub jalankansql(ByVal sQl As String)
        Dim objcmd As New System.Data.OleDb.OleDbCommand
        Call koneksi()
        Try
            objcmd.Connection = con
            objcmd.CommandType = CommandType.Text
            objcmd.CommandText = sQl
            objcmd.ExecuteNonQuery()
            objcmd.Dispose()
            MsgBox("Data sudah disimpan", vbInformation) 'informasi apabila kondisi terpenuhi
        Catch ex As Exception
            MsgBox("Tidak bisa menyimpan data" & ex.Message) 'informasi ketika kondisi tidak terpenuhi
        End Try
    End Sub
    'Untuk mengisi grid
    Sub isigrid(ByVal query3 As String, ByVal nm_table As String, ByVal gried As DataGridView)
        koneksi() 'koneksi ke databse dengan memanggil fungsi koneksi yang ada di module
        da = New OleDb.OleDbDataAdapter(query3, con)
        ds = New DataSet
        ds.Clear()
        da.Fill(ds, nm_table)
        gried.DataSource = (ds.Tables(nm_table))
        gried.Enabled = True
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim a As Boolean
        Dim b As String
        ' Dim c As String
        Dim d As String
        Dim g As String
        Dim h As DataGridView
        h = DataGridView1
        b = "INSERT INTO detail_perawatan (kode_perawatan, kode_tindakan) VALUES('" & TextBox7.Text & "','" & ComboBox1.Text & "')"
        a = TextBox6.Text = ""
        If a Then
            MsgBox("Isi semua data terlebih dahulu!")
        End If
        'c = "UPDATE Matapelajaran SET " + "Nama_Mata_Pelajaran= '" & nm_matpel.Text & "'," + "kkm = '" & kkm.Text & "' WHERE Kode_Mata_Pelajaran = '" & kd_matpel.Text & "'"
        d = "SELECT * FROM detail_perawatan"
        g = "detail_perawatan"
        tblsimpan(a, b, d, g, h)

        Button3.Enabled = True
        Button4.Enabled = True

    End Sub

    Sub tampil5() 'tampil data di griedview
        Call koneksi()

        Dim sql As String = "select tbl_tindakan.kode_tindakan, tbl_tindakan.nama_tindakan, tbl_tindakan.biaya_tindakan From tbl_tindakan right join detail_perawatan ON tbl_tindakan.kode_tindakan=detail_perawatan.kode_tindakan Where detail_perawatan.kode_perawatan='" & TextBox7.Text & "'"
        Dim sqlCommand As New OleDbCommand
        Dim sqlAdapter As New OleDbDataAdapter
        Dim tbl As New DataTable
        With sqlCommand
            .CommandText = sql
            .Connection = con
        End With
        With sqlAdapter
            .SelectCommand = sqlCommand
            .Fill(tbl)
        End With
        DataGridView1.Rows.Clear()
        For i = 0 To tbl.Rows.Count - 1
            With DataGridView1
                .Rows.Add(tbl.Rows(i)("kode_tindakan"), tbl.Rows(i)("nama_tindakan"), tbl.Rows(i)("biaya_tindakan"))
            End With
            DataGridView1.Rows(i).Cells(0).Value = CStr(DataGridView1.RowCount - 1)
            DataGridView1.Rows(i).Cells(1).Value = tbl.Rows(i)("kode_tindakan")
            DataGridView1.Rows(i).Cells(2).Value = tbl.Rows(i)("nama_tindakan")
            DataGridView1.Rows(i).Cells(3).Value = tbl.Rows(i)("biaya_tindakan")
        Next
        hitung()
    End Sub
    Sub hitung()
        Dim totalbiaya As Long
        totalbiaya = 0
        For t As Integer = 0 To DataGridView1.Rows.Count - 1
            totalbiaya = totalbiaya + Val(DataGridView1.Rows(t).Cells(3).Value)
        Next
        tot.Text = totalbiaya
    End Sub

    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call koneksi()
        cmd = New OleDb.OleDbCommand("select kode_perawatan from dt_perawatan order by kode_perawatan desc", con)
        rd = cmd.ExecuteReader
        rd.Read()
        If Not rd.HasRows Then
            TextBox7.Text = "PRW" + "01 "
        Else
            TextBox7.Text = "PRW" + Format(Microsoft.VisualBasic.Right(rd.Item("kode_perawatan"), 2) + 1, "00")

        End If


    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim a As Boolean
        Dim b As String
        ' Dim c As String
        Dim d As String
        Dim g As String
        Dim h As DataGridView
        h = DataGridView1
        b = "INSERT INTO dt_perawatan (kode_perawatan, kode_pasien, total_biaya) VALUES('" & TextBox7.Text & "','" & TextBox1.Text & "','" & tot.Text & "')"
        a = TextBox6.Text = ""
        If a Then
            MsgBox("Isi semua data terlebih dahulu!")
        End If
        'c = "UPDATE Matapelajaran SET " + "Nama_Mata_Pelajaran= '" & nm_matpel.Text & "'," + "kkm = '" & kkm.Text & "' WHERE Kode_Mata_Pelajaran = '" & kd_matpel.Text & "'"
        d = "SELECT * FROM detail_perawatan"
        g = "detail_perawatan"
        tblsimpan(a, b, d, g, h)


    End Sub
End Class
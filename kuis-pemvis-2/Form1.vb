Imports System.Data.OleDb
Public Class Form1
    Dim conn As OleDbConnection
    Dim da As OleDbDataAdapter
    Dim ds As DataSet
    Dim lokasidb As String
    Sub koneksi()
        lokasidb = "Provider = Microsoft.ACE.OLEDB.12.0;DATA Source = pertemuan11.accdb"
        conn = New OleDbConnection(lokasidb)
        If conn.State = ConnectionState.Closed Then conn.Open()

    End Sub
    Sub loaddata()
        da = New OleDbDataAdapter("Select * from mhs", conn)
        ds = New DataSet
        ds.Clear()
        da.Fill(ds, "mhs")
        DataGridView1.DataSource = ds.Tables("mhs")
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        koneksi()
        loaddata()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or ComboBox1.Text = "" Then
            MsgBox("Ada kolom yang belum diisi")
            TextBox1.Focus()
        Else


            Dim dan As DialogResult = MessageBox.Show("Data Berhasil Disimpan!", "pesan", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If dan = DialogResult.No Then

                TextBox1.Text = ""
                TextBox2.Text = ""
                TextBox3.Text = ""
                ComboBox1.Text = ""
            ElseIf dan = DialogResult.Yes Then


                Dim ob As OleDbCommand

                koneksi()
                Dim simpan As String = "insert into mhs values('" & TextBox1.Text & "', '" & TextBox2.Text & "', '" & ComboBox1.Text & "', '" & TextBox3.Text & "')"
                ob = New OleDbCommand(simpan, conn)
                ob.ExecuteNonQuery()

                loaddata()
            End If
        End If


    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click

        Dim ob As OleDbCommand

        koneksi()
        Dim edit As String = "Update mhs SET nama='" & TextBox2.Text & "', jurusan='" & ComboBox1.Text & "', alamat='" & TextBox3.Text & "' where nrp='" & TextBox1.Text & "'"
        ob = New OleDbCommand(edit, conn)
        ob.ExecuteNonQuery()
        MsgBox("Data Berhasil Di Update")

        loaddata()
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        TextBox1.MaxLength = 8
        If e.KeyChar = Chr(13) Then
            koneksi()
            Dim ob As OleDbCommand
            Dim rd As OleDbDataReader
            ob = New OleDbCommand("select * from mhs where nrp='" & TextBox1.Text & "'", conn)
            rd = ob.ExecuteReader
            rd.Read()

            If Not rd.HasRows Then
                MsgBox("Kode Barang Tidak Ada, Silahkan coba lagi!")
                TextBox1.Focus()

            Else
                TextBox2.Text = rd.Item("nama")
                ComboBox1.Text = rd.Item("jurusan")
                TextBox3.Text = rd.Item("alamat")
                TextBox2.Focus()
            End If
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Then
            MsgBox("Silahkan Pilih Data yang akan di hapus dengan Masukan NIM dan ENTER")
        Else
            If MessageBox.Show("Yakin akan dihapus..?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                Call koneksi()
                Dim ob As OleDbCommand
                Dim hapus As String = "delete From mhs where nrp='" & TextBox1.Text & "'"
                ob = New OleDbCommand(hapus, conn)
                ob.ExecuteNonQuery()
            End If
        End If
    End Sub
End Class

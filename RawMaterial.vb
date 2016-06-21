Imports MySql.Data.MySqlClient

Public Class RawMaterial
    Dim constring As String = "Database=inventory;Data Source=localhost;User Id=root;Password="
    Dim conn As New MySqlConnection(constring)
    Sub InLoad()
        Try
            conn.Open()
            Dim stm As String = "SELECT * FROM raw_material"
            Dim DA As New MySqlDataAdapter(stm, conn)
            Dim DS As New DataSet
            DS.Clear()
            DA.Fill(DS, "Raw_Material")
            DataGridView1.DataSource = DS.Tables("Raw_Material")
        Catch ex As MySqlException
            MsgBox("Error: " & ex.ToString())
        Finally
            conn.Close()
        End Try
    End Sub

    Sub Clear()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
    End Sub

    Sub Insert()
        Dim name As String = TextBox2.Text
        Dim stock As Integer = Integer.Parse(TextBox3.Text)
        Try
            conn.Open()
            Dim cmd As New MySqlCommand()
            cmd.Connection = conn
            cmd.CommandText = "INSERT INTO raw_material(itemname, stock) VALUES ('" & name & "', " & stock & ")"
            cmd.Prepare()
            cmd.ExecuteNonQuery()
            MsgBox("Insert berhasil!")
            Clear()
            conn.Close()
        Catch ex As MySqlException
            MsgBox("Insert gagal!")
            conn.Close()
        End Try
    End Sub

    Sub Update1()
        Dim ID As Integer = TextBox1.Text
        Dim name As String = TextBox2.Text
        Dim stock As Integer = Integer.Parse(TextBox3.Text)
        Try
            conn.Open()
            Dim cmd As New MySqlCommand()
            cmd.Connection = conn
            cmd.CommandText = "UPDATE raw_material SET itemname = '" & name & "', stock = " & stock & " WHERE itemID = " & ID
            cmd.Prepare()
            cmd.ExecuteNonQuery()
            MsgBox("Update berhasil!")
            conn.Close()
        Catch ex As MySqlException
            MsgBox("Update gagal!")
            conn.Close()
        End Try
    End Sub

    Sub Delete()
        Dim ID As Integer = 0
        Dim nama As String = TextBox2.Text
        Try
            Dim cmd As MySqlCommand
            Dim DR As MySqlDataReader
            conn.Open()
            Dim query As String = "SELECT itemID FROM raw_material WHERE itemName = '" & nama & "'"
            cmd = New MySqlCommand(query, conn)
            DR = cmd.ExecuteReader()
            DR.Read()
            If DR.HasRows() Then
                ID = DR(0)
            End If
        Catch ex As MySqlException
            MsgBox("Error: " & ex.ToString())
        Finally
            conn.Close()
        End Try
        Try
            conn.Open()
            Dim cmd As New MySqlCommand()
            cmd.Connection = conn
            cmd.CommandText = "DELETE FROM raw_material WHERE itemID = " & ID
            cmd.Prepare()
            cmd.ExecuteNonQuery()
            MsgBox("Delete berhasil!")
            Clear()
            conn.Close()
        Catch ex As MySqlException
            MsgBox("Delete gagal! Barang telah digunakan pada transaksi Incoming / Outgoing!")
            conn.Close()
        End Try
    End Sub

    Private Sub RawMaterial_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        InLoad()
    End Sub

    Private Sub InsertBtn_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub UpdateBtn_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub DeleteBtn_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim arr(3) As String

        For i As Integer = 0 To 2
            If (IsDBNull(DataGridView1.Item(i, DataGridView1.CurrentCell.RowIndex).Value)) Then
                arr(i) = ""
            Else : arr(i) = DataGridView1.Item(i, DataGridView1.CurrentCell.RowIndex).Value.ToString()
            End If
        Next

        TextBox1.Text = arr(0)
        TextBox2.Text = arr(1)
        TextBox3.Text = arr(2)
    End Sub

    Private Sub SearchBtn_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub RefreshBtn_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub MetroButton1_Click(sender As Object, e As EventArgs) Handles MetroButton1.Click
        If TextBox1.Text <> "" Then
            MsgBox("Mohon lakukan refresh terlebih dahulu!")
        Else
            Insert()
        End If
    End Sub

    Private Sub MetroButton2_Click(sender As Object, e As EventArgs) Handles MetroButton2.Click
        If TextBox1.Text = "" Then
            MsgBox("Pilih salah satu data barang terlebih dahulu!")
        Else
            Update1()
        End If
    End Sub

    Private Sub MetroButton3_Click(sender As Object, e As EventArgs) Handles MetroButton3.Click
        If TextBox1.Text = "" Then
            MsgBox("Pilih salah satu data barang terlebih dahulu!")
        Else
            Delete()
        End If

    End Sub

    Private Sub MetroButton4_Click(sender As Object, e As EventArgs) Handles MetroButton4.Click
        Dim found As Boolean = False
        Dim rowcount As Integer = DataGridView1.RowCount
        For i As Integer = 0 To rowcount
            If DataGridView1.Rows(i).Cells(1).Value = TextBox4.Text Then
                DataGridView1.CurrentCell = DataGridView1.Rows(i).Cells(1)
                found = True
                Exit For
            End If
        Next
        If Not found Then
            MsgBox("Nama Item tidak ditemukan")
        End If
        found = False
        TextBox4.Clear()
    End Sub

    Private Sub MetroButton5_Click(sender As Object, e As EventArgs) Handles MetroButton5.Click
        InLoad()
        Clear()
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs)

    End Sub


    Private Sub MetroGrid1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)


    End Sub


    Private Sub DataGridView1_CellContentClick_1(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class
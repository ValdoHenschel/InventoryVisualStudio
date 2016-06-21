Imports MySql.Data.MySqlClient

Public Class Supplier
    Dim constring As String = "Database=inventory;Data Source=localhost;User Id=root;Password="
    Dim conn As New MySqlConnection(constring)

    Sub InLoad()
        Try
            conn.Open()
            Dim stm As String = "SELECT * FROM supplier"
            Dim DA As New MySqlDataAdapter(stm, conn)
            Dim DS As New DataSet
            DS.Clear()
            DA.Fill(DS, "Supplier")
            DataGridView1.DataSource = DS.Tables("Supplier")
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
        TextBox4.Clear()
        TextBox6.Clear()
    End Sub

    Sub Insert()
        Dim name As String = TextBox1.Text
        Dim email As String = TextBox2.Text
        Dim contact As String = TextBox3.Text
        Dim address As String = TextBox4.Text
        Try
            conn.Open()
            Dim cmd As New MySqlCommand()
            cmd.Connection = conn
            cmd.CommandText = "INSERT INTO supplier(supplierName, email, contact, address) VALUES ('" & name & "', '" & email & "', '" & contact & "', '" & address & "')"
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
        Dim name As String = TextBox1.Text
        Dim email As String = TextBox2.Text
        Dim contact As String = TextBox3.Text
        Dim address As String = TextBox4.Text
        Dim ID As Integer = TextBox6.Text
        Try
            conn.Open()
            Dim cmd As New MySqlCommand()
            cmd.Connection = conn
            cmd.CommandText = "UPDATE supplier SET supplierName = '" & name & "', email = '" & email & "', contact = '" & contact & "', address = '" & address & "' WHERE supplierID = " & ID
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
        Dim nama As String = TextBox1.Text
        Try
            Dim cmd As MySqlCommand
            Dim DR As MySqlDataReader
            conn.Open()
            Dim query As String = "SELECT supplierID FROM supplier WHERE supplierName = '" & nama & "'"
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
            cmd.CommandText = "DELETE FROM supplier WHERE supplierID = " & ID
            cmd.Prepare()
            cmd.ExecuteNonQuery()
            MsgBox("Delete berhasil!")
            Clear()
            conn.Close()
        Catch ex As MySqlException
            MsgBox("Delete gagal! Supplier memiliki transaksi Incoming / Outgoing!")
            conn.Close()
        End Try
    End Sub

    Private Sub Supplier_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        InLoad()
    End Sub

    Private Sub InsertBtn_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub UpdateBtn_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub DeleteBtn_Click(sender As Object, e As EventArgs)
        If TextBox6.Text = "" Then
            MsgBox("Pilih salah satu data supplier terlebih dahulu!")
        Else
            Delete()
        End If
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim arr(5) As String

        For i As Integer = 0 To 4
            If (IsDBNull(DataGridView1.Item(i, DataGridView1.CurrentCell.RowIndex).Value)) Then
                arr(i) = ""
            Else : arr(i) = DataGridView1.Item(i, DataGridView1.CurrentCell.RowIndex).Value.ToString()
            End If
        Next

        TextBox6.Text = arr(0)
        TextBox1.Text = arr(1)
        TextBox2.Text = arr(2)
        TextBox3.Text = arr(3)
        TextBox4.Text = arr(4)
    End Sub

    Private Sub SearchBtn_Click(sender As Object, e As EventArgs)
        Dim found As Boolean = False
        Dim rowcount As Integer = DataGridView1.RowCount
        For i As Integer = 0 To rowcount
            If DataGridView1.Rows(i).Cells(1).Value = TextBox5.Text Then
                DataGridView1.CurrentCell = DataGridView1.Rows(i).Cells(1)
                found = True
                Exit For
            End If
        Next
        If Not found Then
            MsgBox("Nama Supplier tidak ditemukan")
        End If
        found = False
        TextBox5.Clear()
    End Sub

    Private Sub RefreshBtn_Click(sender As Object, e As EventArgs)
        InLoad()
        Clear()
    End Sub

    Private Sub Panel3_Paint(sender As Object, e As PaintEventArgs)

    End Sub

    Private Sub MetroButton1_Click(sender As Object, e As EventArgs) Handles MetroButton1.Click
        If TextBox6.Text <> "" Then
            MsgBox("Mohon lakukan refresh terlebih dahulu!")
        Else
            Insert()
        End If
    End Sub

    Private Sub MetroButton2_Click(sender As Object, e As EventArgs) Handles MetroButton2.Click
        If TextBox6.Text = "" Then
            MsgBox("Pilih salah satu data supplier terlebih dahulu!")
        Else
            Update1()
        End If
    End Sub

    Private Sub MetroButton3_Click(sender As Object, e As EventArgs) Handles MetroButton3.Click
        If TextBox6.Text = "" Then
            MsgBox("Pilih salah satu data supplier terlebih dahulu!")
        Else
            Delete()
        End If
    End Sub

    Private Sub MetroButton4_Click(sender As Object, e As EventArgs) Handles MetroButton4.Click
        Dim found As Boolean = False
        Dim rowcount As Integer = DataGridView1.RowCount
        For i As Integer = 0 To rowcount
            If DataGridView1.Rows(i).Cells(1).Value = TextBox5.Text Then
                DataGridView1.CurrentCell = DataGridView1.Rows(i).Cells(1)
                found = True
                Exit For
            End If
        Next
        If Not found Then
            MsgBox("Nama Supplier tidak ditemukan")
        End If
        found = False
        TextBox5.Clear()
    End Sub

    Private Sub MetroButton5_Click(sender As Object, e As EventArgs) Handles MetroButton5.Click
        InLoad()
        Clear()
    End Sub
End Class
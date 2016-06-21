Imports MySql.Data.MySqlClient

Public Class CustomerField
    Dim constring As String = "Database=inventory;Data Source=localhost;User Id=root;Password="
    Dim conn As New MySqlConnection(constring)

    Sub InLoad()
        Try
            conn.Open()
            Dim stm As String = "SELECT * FROM customer"
            Dim DA As New MySqlDataAdapter(stm, conn)
            Dim DS As New DataSet
            DS.Clear()
            DA.Fill(DS, "Customer")
            DataGridView1.DataSource = DS.Tables("Customer")
        Catch ex As MySqlException
            MsgBox("Error: " & ex.ToString())
        Finally
            conn.Close()
        End Try
    End Sub

    Sub Clear()
        MetroTextBox1.Clear()
        MetroTextBox2.Clear()
        MetroTextBox3.Clear()
        MetroTextBox4.Clear()
        MetroTextBox6.Clear()
    End Sub

    Sub Insert()
        Dim name As String = MetroTextBox1.Text
        Dim email As String = MetroTextBox2.Text
        Dim contact As String = MetroTextBox3.Text
        Dim address As String = MetroTextBox4.Text
        Try
            conn.Open()
            Dim cmd As New MySqlCommand()
            cmd.Connection = conn
            cmd.CommandText = "INSERT INTO customer(name, email, contact, address) VALUES ('" & name & "', '" & email & "', '" & contact & "', '" & address & "')"
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
        Dim name As String = MetroTextBox1.Text
        Dim email As String = MetroTextBox2.Text
        Dim contact As String = MetroTextBox3.Text
        Dim address As String = MetroTextBox4.Text
        Dim ID As Integer = MetroTextBox6.Text
        Try
            conn.Open()
            Dim cmd As New MySqlCommand()
            cmd.Connection = conn
            cmd.CommandText = "UPDATE customer SET name = '" & name & "', email = '" & email & "', contact = '" & contact & "', address = '" & address & "' WHERE customerID = " & ID
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
        Dim nama As String = MetroTextBox1.Text
        Try
            Dim cmd As MySqlCommand
            Dim DR As MySqlDataReader
            conn.Open()
            Dim query As String = "SELECT customerID FROM customer WHERE name = '" & nama & "'"
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
            cmd.CommandText = "DELETE FROM customer WHERE customerID = " & ID
            cmd.Prepare()
            cmd.ExecuteNonQuery()
            MsgBox("Delete berhasil!")
            Clear()
            conn.Close()
        Catch ex As MySqlException
            MsgBox("Delete gagal! Customer memiliki transaksi Finished Goods!")
            conn.Close()
        End Try
    End Sub

    Private Sub CustomerField_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        InLoad()
    End Sub

    Private Sub InsertBtn_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub UpdateBtn_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub DeleteBtn_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub MetroDataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs)
        Dim arr(5) As String

        For i As Integer = 0 To 4
            If (IsDBNull(DataGridView1.Item(i, DataGridView1.CurrentCell.RowIndex).Value)) Then
                arr(i) = ""
            Else : arr(i) = DataGridView1.Item(i, DataGridView1.CurrentCell.RowIndex).Value.ToString()
            End If
        Next

        MetroTextBox6.Text = arr(0)
        MetroTextBox1.Text = arr(1)
        MetroTextBox2.Text = arr(2)
        MetroTextBox3.Text = arr(3)
        MetroTextBox4.Text = arr(4)
    End Sub

    Private Sub SearchBtn_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Refresh_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub MetroButton1_Click(sender As Object, e As EventArgs) Handles MetroButton1.Click
        If MetroTextBox6.Text <> "" Then
            MsgBox("Mohon lakukan refresh terlebih dahulu!")
        Else
            Insert()
        End If
    End Sub

    Private Sub MetroButton2_Click(sender As Object, e As EventArgs) Handles MetroButton2.Click
        If MetroTextBox6.Text = "" Then
            MsgBox("Pilih salah satu data customer terlebih dahulu!")
        Else
            Update1()
        End If
    End Sub

    Private Sub MetroButton3_Click(sender As Object, e As EventArgs) Handles MetroButton3.Click
        If MetroTextBox6.Text = "" Then
            MsgBox("Pilih salah satu data customer terlebih dahulu!")
        Else
            Delete()
        End If
    End Sub

    Private Sub MetroLabel4_Click(sender As Object, e As EventArgs) Handles MetroLabel4.Click, MetroLabel5.Click

    End Sub

    Private Sub MetroButton4_Click(sender As Object, e As EventArgs) Handles MetroButton4.Click
        Dim found As Boolean = False
        Dim rowcount As Integer = DataGridView1.RowCount
        For i As Integer = 0 To rowcount
            If DataGridView1.Rows(i).Cells(1).Value = MetroTextBox5.Text Then
                DataGridView1.CurrentCell = DataGridView1.Rows(i).Cells(1)
                found = True
                Exit For
            End If
        Next
        If Not found Then
            MsgBox("Nama Customer tidak ditemukan")
        End If
        found = False
        MetroTextBox5.Clear()
    End Sub

    Private Sub MetroButton5_Click(sender As Object, e As EventArgs) Handles MetroButton5.Click
        InLoad()
        Clear()
    End Sub
    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim arr(5) As String

        For i As Integer = 0 To 4
            If (IsDBNull(DataGridView1.Item(i, DataGridView1.CurrentCell.RowIndex).Value)) Then
                arr(i) = ""
            Else : arr(i) = DataGridView1.Item(i, DataGridView1.CurrentCell.RowIndex).Value.ToString()
            End If
        Next

        MetroTextBox6.Text = arr(0)
        MetroTextBox1.Text = arr(1)
        MetroTextBox2.Text = arr(2)
        MetroTextBox3.Text = arr(3)
        MetroTextBox4.Text = arr(4)
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Panel3_Paint(sender As Object, e As PaintEventArgs)

    End Sub

    Private Sub MetroTextBox5_Click(sender As Object, e As EventArgs) Handles MetroTextBox5.Click

    End Sub
End Class
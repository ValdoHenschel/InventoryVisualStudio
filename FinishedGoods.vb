Imports MySql.Data.MySqlClient

Public Class FinishedGoods
    Dim constring = "Database=inventory;Data Source=localhost;User Id=root;Password="
    Dim conn As New MySqlConnection(constring)

    Sub InLoad()
        Try
            conn.Open()
            Dim stm As String = "SELECT a.goodsID ""No."", b.name ""Nama Pelanggan"", a.finishedDate ""Tanggal Selesai"", a.takenDate ""Tanggal Diambil"" FROM finished_goods a JOIN customer b ON a.customerID = b.customerID ORDER BY 1"
            Dim DA As New MySqlDataAdapter(stm, conn)
            Dim DS As New DataSet
            DS.Clear()
            DA.Fill(DS, "Finished")
            DataGridView1.DataSource = DS.Tables("Finished")
        Catch ex As MySqlException
            MsgBox("Error: " & ex.ToString())
        Finally
            conn.Close()
        End Try
    End Sub

    Sub InitLoad()
        Dim cmd As MySqlCommand
        Dim DR As MySqlDataReader
        ComboBox1.Items.Clear()
        Try
            conn.Open()
            Dim query As String = "SELECT name FROM customer ORDER BY 1"
            cmd = New MySqlCommand(query, conn)
            DR = cmd.ExecuteReader()
            While (DR.Read())
                ComboBox1.Items.Add(DR(0))
            End While
        Catch ex As MySqlException
            MsgBox("Error: " & ex.ToString())
        Finally
            conn.Close()
        End Try
    End Sub

    Sub Clear()
        TextBox1.Clear()
        ComboBox1.SelectedIndex = -1
        ComboBox1.Text = String.Empty
    End Sub

    Public Sub FinishedGoods_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        InLoad()
        InitLoad()
    End Sub
    Private Sub SearchBtn_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub RefreshBtn_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim arr(3) As String
        For i As Integer = 0 To 3
            If (IsDBNull(DataGridView1.Item(i, DataGridView1.CurrentCell.RowIndex).Value)) Then
                arr(i) = ""
            Else : arr(i) = DataGridView1.Item(i, DataGridView1.CurrentCell.RowIndex).Value.ToString()
            End If
        Next
        TextBox1.Text = arr(0)
        Dim temp As String
        For i As Integer = 0 To ComboBox1.Items.Count
            ComboBox1.SelectedIndex = i
            temp = ComboBox1.Text
            If temp = arr(1) Then
                ComboBox1.SelectedIndex = i
                Exit For
            End If
        Next
    End Sub

    Private Sub MetroButton1_Click(sender As Object, e As EventArgs) Handles MetroButton1.Click
        If TextBox1.Text <> "" Then
            MsgBox("Mohon lakukan refresh terlebih dahulu!")
        Else
            Dim customer As String = ComboBox1.Text
            Dim customerID As Integer
            Dim tanggal = Date.Now.ToString("YYYY-MM-DD")
            Try
                conn.Open()
                Dim cmd As MySqlCommand
                Dim DR As MySqlDataReader
                Dim query As String = "SELECT customerID FROM customer WHERE name = '" & customer & "'"
                cmd = New MySqlCommand(query, conn)
                DR = cmd.ExecuteReader
                DR.Read()
                If DR.HasRows Then
                    customerID = DR(0)
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
                cmd.CommandText = "INSERT INTO finished_goods (customerID, finishedDate) VALUES (" & customerID & ", CURDATE())"
                cmd.Prepare()
                cmd.ExecuteNonQuery()
                MsgBox("Insert berhasil!")
                Clear()
                conn.Close()
            Catch ex As MySqlException
                MsgBox("Insert gagal!")
                conn.Close()
            End Try
        End If
    End Sub

    Private Sub MetroButton2_Click(sender As Object, e As EventArgs) Handles MetroButton2.Click
        If TextBox1.Text = "" Then
            MsgBox("Pilih salah satu transaksi terlebih dahulu!")
        Else
            Dim finishedID As Integer = TextBox1.Text
            Dim customer As String = ComboBox1.Text
            Dim customerID As Integer
            DateTimePicker1.Format = DateTimePickerFormat.Custom
            DateTimePicker1.CustomFormat = "yyyy-MM-dd"
            Dim finishedDate As String = DateTimePicker1.Text
            MetroDateTime1.Format = DateTimePickerFormat.Custom
            MetroDateTime1.CustomFormat = "yyyy-MM-dd"
            Dim takenDate As String = MetroDateTime1.Text
            Try
                conn.Open()
                Dim cmd As MySqlCommand
                Dim DR As MySqlDataReader
                Dim query As String = "SELECT customerID FROM customer WHERE name = '" & customer & "'"
                cmd = New MySqlCommand(query, conn)
                DR = cmd.ExecuteReader()
                DR.Read()
                If DR.HasRows Then
                    customerID = DR(0)
                End If
            Catch ex As MySqlException
                MsgBox("Error: " & ex.ToString())
            Finally
                conn.Close()
            End Try
            Try
                conn.Open()
                Dim cmd As New MySqlCommand
                cmd.Connection = conn
                cmd.CommandText = "UPDATE finished_goods SET customerID = " & customerID & ", finishedDate = '" & finishedDate & "', takenDate = '" & takenDate & "' WHERE goodsID = " & finishedID
                cmd.Prepare()
                cmd.ExecuteNonQuery()
                MsgBox("Update berhasil!")
                Clear()
                conn.Close()
            Catch ex As MySqlException
                MsgBox("Update gagal!")
                conn.Close()
            End Try
            DateTimePicker1.Format = DateTimePickerFormat.Long
        End If
    End Sub

    Private Sub MetroButton3_Click(sender As Object, e As EventArgs) Handles MetroButton3.Click
        If TextBox1.Text = "" Then
            MsgBox("Pilih salah satu data transaksi terlebih dahulu!")
        Else
            Dim finishedID As Integer = TextBox1.Text
            Try
                conn.Open()
                Dim cmd As New MySqlCommand()
                cmd.Connection = conn
                cmd.CommandText = "DELETE FROM finished_goods WHERE goodsID = " & finishedID
                cmd.Prepare()
                cmd.ExecuteNonQuery()
                MsgBox("Delete berhasil!")
                Clear()
                conn.Close()
            Catch ex As MySqlException
                MsgBox("Delete gagal!")
                conn.Close()
            End Try
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub

 
    Private Sub MetroButton4_Click(sender As Object, e As EventArgs) Handles MetroButton4.Click
        If IsDBNull(TextBox3.Text) Then
        Else
            Dim found As Boolean = False
            Dim rowcount As Integer = DataGridView1.RowCount
            For i As Integer = 0 To rowcount
                If DataGridView1.Rows(i).Cells(1).Value = TextBox3.Text Then
                    DataGridView1.CurrentCell = DataGridView1.Rows(i).Cells(1)
                    found = True
                    Exit For
                End If
            Next
            If Not found Then
                MsgBox("Keyword tidak ditemukan!")
            End If
            found = False
            TextBox3.Clear()
        End If
    End Sub

    Private Sub MetroButton5_Click(sender As Object, e As EventArgs) Handles MetroButton5.Click
        InLoad()
        InitLoad()
        Clear()
    End Sub
    Private Sub SetBorderAndGridlineStyles()

        With Me.DataGridView1
            .GridColor = Color.BlueViolet
            .BorderStyle = Windows.Forms.BorderStyle.Fixed3D
            .CellBorderStyle = DataGridViewCellBorderStyle.None
            .RowHeadersBorderStyle = _
                DataGridViewHeaderBorderStyle.Single
            .ColumnHeadersBorderStyle = _
                DataGridViewHeaderBorderStyle.Single
        End With

    End Sub

    Private Sub DataGridView1_CellContentClick_1(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class
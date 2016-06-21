﻿Imports MySql.Data.MySqlClient

Public Class Outgoing
    Dim constring = "Database=inventory;Data source=localhost;User Id=root;Password="
    Dim conn As New MySqlConnection(constring)

    Sub InLoad()
        Try
            conn.Open()
            Dim stm As String = "SELECT outgoingID ""ID Pemesanan"", outgoingDate ""Tanggal Pemesanan"" FROM outgoing_item ORDER BY 1"
            Dim DA As New MySqlDataAdapter(stm, conn)
            Dim DS As New DataSet
            DS.Clear()
            DA.Fill(DS, "Outgoing")
            DataGridView1.DataSource = DS.Tables("Outgoing")
            Dim stm2 As String = "SELECT a.outgoingDetailID ""No."", a.outgoingID ""ID Pemesanan"", b.itemName ""Jenis Barang"" , a.quantity ""Jumlah"" FROM outgoing_detail a JOIN raw_material b ON a.itemID = b.itemID ORDER BY 1"
            Dim DA2 As New MySqlDataAdapter(stm2, conn)
            Dim DS2 As New DataSet
            DS2.Clear()
            DA2.Fill(DS2, "Detail")
            DataGridView2.DataSource = DS2.Tables("Detail")
        Catch ex As MySqlException
            MsgBox("Error: " & ex.ToString())
        Finally
            conn.Close()
        End Try
    End Sub

    Sub InitLoad()
        Dim cmd As MySqlCommand
        Dim DR As MySqlDataReader
        ComboBox2.Items.Clear()
        ComboBox3.Items.Clear()

        Try
            conn.Open()
            Dim query As String = "SELECT itemName FROM raw_material ORDER BY itemName"
            cmd = New MySqlCommand(query, conn)
            DR = cmd.ExecuteReader()
            While (DR.Read())
                ComboBox2.Items.Add(DR(0))
            End While
        Catch ex As MySqlException
            MsgBox("Error: " & ex.ToString())
        Finally
            conn.Close()
        End Try

        Try
            conn.Open()
            Dim query As String = "SELECT outgoingID FROM outgoing_item ORDER BY outgoingID"
            cmd = New MySqlCommand(query, conn)
            DR = cmd.ExecuteReader()
            While (DR.Read())
                ComboBox3.Items.Add(DR(0))
            End While
        Catch ex As MySqlException
            MsgBox("Error: " & ex.ToString())
        Finally
            conn.Close()
        End Try
    End Sub

    Sub Clear()
        TextBox1.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox6.Clear()
        ComboBox2.SelectedIndex = -1
        ComboBox2.Text = String.Empty
        ComboBox2.SelectedIndex = -1
        ComboBox2.Text = String.Empty
        ComboBox3.SelectedIndex = -1
        ComboBox3.Text = String.Empty
    End Sub

    Private Sub Outgoing_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        InLoad()
        InitLoad()
        Clear()
    End Sub

    Private Sub InsertBtn_Click(sender As Object, e As EventArgs)
        If TextBox1.Text <> "" Then
            MsgBox("Mohon lakukan refresh terlebih dahulu!")
        Else
            DateTimePicker1.Format = DateTimePickerFormat.Custom
            DateTimePicker1.CustomFormat = "yyyy-MM-dd"
            Dim outgoingDate As String = DateTimePicker1.Text
            Try
                conn.Open()
                Dim cmd As New MySqlCommand()
                cmd.Connection = conn
                cmd.CommandText = "INSERT INTO outgoing_item (outgoingDate) VALUES ('" & outgoingDate & "')"
                cmd.Prepare()
                cmd.ExecuteNonQuery()
                MsgBox("Insert berhasil!")
                Clear()
                conn.Close()
            Catch ex As MySqlException
                MsgBox("Insert gagal!")
                conn.Close()
            End Try
            DateTimePicker1.Format = DateTimePickerFormat.Long
        End If
    End Sub

    Private Sub UpdateBtn_Click(sender As Object, e As EventArgs)
        If TextBox1.Text = "" Then
            MsgBox("Pilih salah satu data transaksi terlebih dahulu!")
        Else
            Dim outgoingID As Integer = TextBox1.Text
            DateTimePicker1.Format = DateTimePickerFormat.Custom
            DateTimePicker1.CustomFormat = "yyyy-MM-dd"
            Dim outgoingDate As String = DateTimePicker1.Text
            Try
                conn.Open()
                Dim cmd As New MySqlCommand()
                cmd.Connection = conn
                cmd.CommandText = "UPDATE outgoing_item SET outgoingDate = '" & outgoingDate & "' WHERE outgoingID = " & outgoingID
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

    Private Sub DeleteBtn_Click(sender As Object, e As EventArgs)
        If TextBox1.Text = "" Then
            MsgBox("Pilih salah satu data transaksi terlebih dahulu!")
        Else
            Dim outgoingID As Integer = TextBox1.Text
            Dim itemID As Integer
            Dim quantity As Integer
            Try
                Dim cmd As MySqlCommand
                Dim DR As MySqlDataReader
                conn.Open()
                Dim query As String = "SELECT itemID, quantity FROM outgoing_detail WHERE outgoingID = " & outgoingID
                cmd = New MySqlCommand(query, conn)
                DR = cmd.ExecuteReader()
                If (DR.HasRows) Then
                    While (DR.Read())
                        Dim conn2 As New MySqlConnection(constring)
                        conn2.Open()
                        Dim cmd2 As New MySqlCommand()
                        cmd2.Connection = conn2
                        itemID = DR(0)
                        quantity = DR(1)
                        cmd2.CommandText = "UPDATE raw_material SET stock = stock + " & quantity & " WHERE itemID = " & itemID
                        cmd2.Prepare()
                        cmd2.ExecuteNonQuery()
                        conn2.Close()
                    End While
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
                cmd.CommandText = "DELETE FROM outgoing_detail WHERE outgoingID = " & outgoingID
                cmd.Prepare()
                cmd.ExecuteNonQuery()
                cmd.CommandText = "DELETE FROM outgoing_item WHERE outgoingID = " & outgoingID
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

    Private Sub SearchBtn_Click(sender As Object, e As EventArgs)
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
    End Sub

    Private Sub RefreshBtn_Click(sender As Object, e As EventArgs)
        InLoad()
        InitLoad()
        Clear()
    End Sub

    Private Sub InsertBtn2_Click(sender As Object, e As EventArgs)
        If TextBox1.Text <> "" Then
            MsgBox("Mohon lakukan refresh terlebih dahulu!")
        Else
            Dim outgoingID As Integer = ComboBox3.Text
            Dim item As String = ComboBox3.Text
            Dim quantity As Integer = TextBox1.Text
            Dim stock As Integer
            Dim itemID As Integer
            Try
                conn.Open()
                Dim cmd As MySqlCommand
                Dim DR As MySqlDataReader
                Dim query As String = "SELECT itemID FROM raw_material WHERE itemName = '" & item & "'"
                cmd = New MySqlCommand(query, conn)
                DR = cmd.ExecuteReader()
                DR.Read()
                If DR.HasRows Then
                    itemID = DR(0)
                End If
            Catch ex As MySqlException
                MsgBox("Error: " & ex.ToString())
            Finally
                conn.Close()
            End Try
            Try
                conn.Open()
                Dim cmd As MySqlCommand
                Dim DR As MySqlDataReader
                Dim query As String = "SELECT stock FROM raw_material WHERE itemID = " & itemID
                cmd = New MySqlCommand(query, conn)
                DR = cmd.ExecuteReader()
                DR.Read()
                If DR.HasRows Then
                    stock = DR(0)
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
                cmd.CommandText = "INSERT INTO outgoing_detail(outgoingID, itemID, quantity) VALUES (" & outgoingID & ", " & itemID & ", " & quantity & ")"
                cmd.Prepare()
                cmd.ExecuteNonQuery()
                stock -= quantity
                cmd.CommandText = "UPDATE raw_material SET stock = " & stock & " WHERE itemID = " & itemID
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

    Private Sub UpdateBtn2_Click(sender As Object, e As EventArgs)
        If TextBox1.Text = "" Then
            MsgBox("Pilih salah satu transaksi terlebih dahulu!")
        Else
            Dim outgoingDetailID As Integer = TextBox1.Text
            Dim outgoingID As Integer = ComboBox3.Text
            Dim item As String = ComboBox3.Text
            Dim quantity As Integer = TextBox1.Text

            Dim itemID As Integer
            Dim initialQuantity As Integer
            Dim initialItem As Integer
            Dim initialStock As Integer
            Dim stock As Integer

            Try
                'Ambil itemID di textbox, terus ambil stoknya
                conn.Open()
                Dim cmd As MySqlCommand
                Dim DR As MySqlDataReader
                Dim query As String = "SELECT itemID, stock FROM raw_material WHERE itemName = '" & item & "'"
                cmd = New MySqlCommand(query, conn)
                DR = cmd.ExecuteReader()
                DR.Read()
                If DR.HasRows Then
                    'Dapet itemID sama stok yang mau di update
                    itemID = DR(0)
                    stock = DR(1)
                End If
            Catch ex As MySqlException
                MsgBox("Error: " & ex.ToString())
            Finally
                conn.Close()
            End Try

            Try
                'Ambil data transaksi sebelumnya
                conn.Open()
                Dim cmd As MySqlCommand
                Dim DR As MySqlDataReader
                Dim query As String = "SELECT quantity, itemID FROM outgoing_detail WHERE outgoingDetailID = " & outgoingDetailID
                cmd = New MySqlCommand(query, conn)
                DR = cmd.ExecuteReader()
                DR.Read()
                If DR.HasRows Then
                    'Dapet itemID sama jumlah quantity transaksi lama
                    initialQuantity = DR(0)
                    initialItem = DR(1)
                End If
            Catch ex As MySqlException
                MsgBox("Error: " & ex.ToString())
            Finally
                conn.Close()
            End Try

            Try
                'Ambil data stok lama dari transaksi awal
                conn.Open()
                Dim cmd As MySqlCommand
                Dim DR As MySqlDataReader
                Dim query As String = "SELECT stock FROM raw_material WHERE itemID = " & initialItem
                cmd = New MySqlCommand(query, conn)
                DR = cmd.ExecuteReader()
                DR.Read()
                If DR.HasRows Then
                    'Dapet stok dari transaksi lama
                    initialStock = DR(0)
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
                cmd.CommandText = "UPDATE outgoing_detail SET itemID = " & itemID & ", quantity = " & quantity & " WHERE outgoingDetailID = " & outgoingDetailID
                cmd.Prepare()
                cmd.ExecuteNonQuery()
                If initialItem <> itemID Then
                    Dim jumlahlama As Integer = initialStock + initialQuantity
                    cmd.CommandText = "UPDATE raw_material SET stock = " & jumlahlama & " WHERE itemID = " & initialItem
                    cmd.Prepare()
                    cmd.ExecuteNonQuery()
                    Dim jumlahbaru2 As Integer = stock - quantity
                    cmd.CommandText = "UPDATE raw_material SET stock = " & jumlahbaru2 & " WHERE itemID = " & itemID
                    cmd.Prepare()
                    cmd.ExecuteNonQuery()
                Else
                    Dim jumlahbaru As Integer = stock + initialQuantity - quantity
                    cmd.CommandText = "UPDATE raw_material SET stock = " & jumlahbaru & " WHERE itemID = " & itemID
                    cmd.Prepare()
                    cmd.ExecuteNonQuery()
                End If
                conn.Close()
                MsgBox("Update berhasil!")
            Catch ex As MySqlException
                MsgBox("Update gagal!")
                conn.Close()
            End Try
            End
            If True Then

            End If
        End If


    End Sub

    Private Sub DeleteBtn2_Click(sender As Object, e As EventArgs)
        If TextBox1.Text = "" Then
            MsgBox("Pilih salah satu transaksi terlebih dahulu!")
        Else
            Dim outgoingDetailID As Integer = TextBox1.Text
            Dim item As String = TextBox1.Text
            Dim quantity As Integer
            Dim itemID As Integer
            Dim stock As Integer
            Try
                conn.Open()
                Dim cmd As MySqlCommand
                Dim DR As MySqlDataReader
                Dim query As String = "SELECT itemID, stock FROM raw_material WHERE itemName = '" & item & "'"
                cmd = New MySqlCommand(query, conn)
                DR = cmd.ExecuteReader()
                DR.Read()
                If DR.HasRows Then
                    itemID = DR(0)
                    stock = DR(1)
                End If
            Catch ex As MySqlException
                MsgBox("Error: " & ex.ToString())
            Finally
                conn.Close()
            End Try
            Try
                conn.Open()
                Dim cmd As MySqlCommand
                Dim DR As MySqlDataReader
                Dim query As String = "SELECT quantity FROM outgoing_detail WHERE outgoingDetailID = " & outgoingDetailID
                cmd = New MySqlCommand(query, conn)
                DR = cmd.ExecuteReader
                DR.Read()
                If DR.HasRows Then
                    quantity = DR(0)
                End If
                conn.Close()
            Catch ex As MySqlException
                MsgBox("Error: " & ex.ToString())
            Finally
                conn.Close()
            End Try
            Try
                conn.Open()
                Dim cmd As New MySqlCommand()
                cmd.Connection = conn
                cmd.CommandText = "DELETE FROM outgoing_detail WHERE outgoingDetailID = " & outgoingDetailID & " AND itemID = " & itemID
                cmd.Prepare()
                cmd.ExecuteNonQuery()
                stock += quantity
                cmd.CommandText = "UPDATE raw_material SET stock = " & stock & " WHERE itemID = " & itemID
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

    Private Sub SearchBtn2_Click(sender As Object, e As EventArgs)
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
            MsgBox("Keyword tidak ditemukan")
        End If
        found = False
        TextBox5.Clear()
    End Sub

    Private Sub RefreshBtn2_Click(sender As Object, e As EventArgs)
        InitLoad()
        InLoad()
        Clear()
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim arr(2) As String
        For i As Integer = 0 To 1
            If (IsDBNull(DataGridView1.Item(i, DataGridView1.CurrentCell.RowIndex).Value)) Then
                arr(i) = ""
            Else : arr(i) = DataGridView1.Item(i, DataGridView1.CurrentCell.RowIndex).Value.ToString()
            End If
        Next
        TextBox1.Text = arr(0)
        Dim temp As String
        For i As Integer = 0 To ComboBox3.Items.Count
            ComboBox3.SelectedIndex = i
            temp = ComboBox3.SelectedItem
            If temp = arr(0) Then
                ComboBox3.SelectedIndex = i
                Exit For
            End If
        Next
        Try
            conn.Open()
            Dim stm2 As String = "SELECT a.outgoingDetailID ""No."", a.outgoingID ""ID Pemesanan"", b.itemName ""Jenis Barang"" , a.quantity ""Jumlah"" FROM outgoing_detail a JOIN raw_material b ON a.itemID = b.itemID WHERE outgoingID = " & arr(0) & " ORDER BY 1"
            Dim DA2 As New MySqlDataAdapter(stm2, conn)
            Dim DS2 As New DataSet
            DS2.Clear()
            DA2.Fill(DS2, "Detail")
            DataGridView2.DataSource = DS2.Tables("Detail")
        Catch ex As MySqlException
            MsgBox("Error: " & ex.ToString())
        Finally
            conn.Close()
        End Try
    End Sub

    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        Dim arr(4) As String
        For i As Integer = 0 To 3
            If (IsDBNull(DataGridView2.Item(i, DataGridView2.CurrentCell.RowIndex).Value)) Then
                arr(i) = ""
            Else : arr(i) = DataGridView2.Item(i, DataGridView2.CurrentCell.RowIndex).Value.ToString()
            End If
        Next
        TextBox4.Text = arr(3)
        TextBox6.Text = arr(0)
        Dim temp1 As String
        For i As Integer = 0 To ComboBox3.Items.Count
            ComboBox3.SelectedIndex = i
            temp1 = ComboBox3.SelectedItem
            If temp1 = arr(1) Then
                ComboBox3.SelectedIndex = i
                Exit For
            End If
        Next
        Dim temp2 As String
        For i As Integer = 0 To ComboBox2.Items.Count
            ComboBox2.SelectedIndex = i
            temp2 = ComboBox2.Text
            If temp2 = arr(2) Then
                ComboBox2.SelectedIndex = i
                Exit For
            End If
        Next
    End Sub

    Private Sub MetroButton1_Click(sender As Object, e As EventArgs) Handles MetroButton1.Click
        If TextBox6.Text <> "" Then
            MsgBox("Mohon lakukan refresh terlebih dahulu!")
        Else
            Dim outgoingID As Integer = ComboBox3.Text
            Dim item As String = ComboBox2.Text
            Dim quantity As Integer = TextBox4.Text
            Dim stock As Integer
            Dim itemID As Integer
            Try
                conn.Open()
                Dim cmd As MySqlCommand
                Dim DR As MySqlDataReader
                Dim query As String = "SELECT itemID FROM raw_material WHERE itemName = '" & item & "'"
                cmd = New MySqlCommand(query, conn)
                DR = cmd.ExecuteReader()
                DR.Read()
                If DR.HasRows Then
                    itemID = DR(0)
                End If
            Catch ex As MySqlException
                MsgBox("Error: " & ex.ToString())
            Finally
                conn.Close()
            End Try
            Try
                conn.Open()
                Dim cmd As MySqlCommand
                Dim DR As MySqlDataReader
                Dim query As String = "SELECT stock FROM raw_material WHERE itemID = " & itemID
                cmd = New MySqlCommand(query, conn)
                DR = cmd.ExecuteReader()
                DR.Read()
                If DR.HasRows Then
                    stock = DR(0)
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
                cmd.CommandText = "INSERT INTO outgoing_detail(outgoingID, itemID, quantity) VALUES (" & outgoingID & ", " & itemID & ", " & quantity & ")"
                cmd.Prepare()
                cmd.ExecuteNonQuery()
                stock -= quantity
                cmd.CommandText = "UPDATE raw_material SET stock = " & stock & " WHERE itemID = " & itemID
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
        If TextBox6.Text = "" Then
            MsgBox("Pilih salah satu transaksi terlebih dahulu!")
        Else
            Dim outgoingDetailID As Integer = TextBox6.Text
            Dim outgoingID As Integer = ComboBox3.Text
            Dim item As String = ComboBox2.Text
            Dim quantity As Integer = TextBox4.Text

            Dim itemID As Integer
            Dim initialQuantity As Integer
            Dim initialItem As Integer
            Dim initialStock As Integer
            Dim stock As Integer

            Try
                'Ambil itemID di textbox, terus ambil stoknya
                conn.Open()
                Dim cmd As MySqlCommand
                Dim DR As MySqlDataReader
                Dim query As String = "SELECT itemID, stock FROM raw_material WHERE itemName = '" & item & "'"
                cmd = New MySqlCommand(query, conn)
                DR = cmd.ExecuteReader()
                DR.Read()
                If DR.HasRows Then
                    'Dapet itemID sama stok yang mau di update
                    itemID = DR(0)
                    stock = DR(1)
                End If
            Catch ex As MySqlException
                MsgBox("Error: " & ex.ToString())
            Finally
                conn.Close()
            End Try

            Try
                'Ambil data transaksi sebelumnya
                conn.Open()
                Dim cmd As MySqlCommand
                Dim DR As MySqlDataReader
                Dim query As String = "SELECT quantity, itemID FROM outgoing_detail WHERE outgoingDetailID = " & outgoingDetailID
                cmd = New MySqlCommand(query, conn)
                DR = cmd.ExecuteReader()
                DR.Read()
                If DR.HasRows Then
                    'Dapet itemID sama jumlah quantity transaksi lama
                    initialQuantity = DR(0)
                    initialItem = DR(1)
                End If
            Catch ex As MySqlException
                MsgBox("Error: " & ex.ToString())
            Finally
                conn.Close()
            End Try

            Try
                'Ambil data stok lama dari transaksi awal
                conn.Open()
                Dim cmd As MySqlCommand
                Dim DR As MySqlDataReader
                Dim query As String = "SELECT stock FROM raw_material WHERE itemID = " & initialItem
                cmd = New MySqlCommand(query, conn)
                DR = cmd.ExecuteReader()
                DR.Read()
                If DR.HasRows Then
                    'Dapet stok dari transaksi lama
                    initialStock = DR(0)
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
                cmd.CommandText = "UPDATE outgoing_detail SET itemID = " & itemID & ", quantity = " & quantity & " WHERE outgoingDetailID = " & outgoingDetailID
                cmd.Prepare()
                cmd.ExecuteNonQuery()
                If initialItem <> itemID Then
                    Dim jumlahlama As Integer = initialStock + initialQuantity
                    cmd.CommandText = "UPDATE raw_material SET stock = " & jumlahlama & " WHERE itemID = " & initialItem
                    cmd.Prepare()
                    cmd.ExecuteNonQuery()
                    Dim jumlahbaru2 As Integer = stock - quantity
                    cmd.CommandText = "UPDATE raw_material SET stock = " & jumlahbaru2 & " WHERE itemID = " & itemID
                    cmd.Prepare()
                    cmd.ExecuteNonQuery()
                Else
                    Dim jumlahbaru As Integer = stock + initialQuantity - quantity
                    cmd.CommandText = "UPDATE raw_material SET stock = " & jumlahbaru & " WHERE itemID = " & itemID
                    cmd.Prepare()
                    cmd.ExecuteNonQuery()
                End If
                conn.Close()
                MsgBox("Update berhasil!")
            Catch ex As MySqlException
                MsgBox("Update gagal!")
                conn.Close()
            End Try
        End If
    End Sub

    Private Sub MetroButton3_Click(sender As Object, e As EventArgs) Handles MetroButton3.Click
        If TextBox6.Text = "" Then
            MsgBox("Pilih salah satu transaksi terlebih dahulu!")
        Else
            Dim outgoingDetailID As Integer = TextBox6.Text
            Dim item As String = ComboBox2.Text
            Dim quantity As Integer
            Dim itemID As Integer
            Dim stock As Integer
            Try
                conn.Open()
                Dim cmd As MySqlCommand
                Dim DR As MySqlDataReader
                Dim query As String = "SELECT itemID, stock FROM raw_material WHERE itemName = '" & item & "'"
                cmd = New MySqlCommand(query, conn)
                DR = cmd.ExecuteReader()
                DR.Read()
                If DR.HasRows Then
                    itemID = DR(0)
                    stock = DR(1)
                End If
            Catch ex As MySqlException
                MsgBox("Error: " & ex.ToString())
            Finally
                conn.Close()
            End Try
            Try
                conn.Open()
                Dim cmd As MySqlCommand
                Dim DR As MySqlDataReader
                Dim query As String = "SELECT quantity FROM outgoing_detail WHERE outgoingDetailID = " & outgoingDetailID
                cmd = New MySqlCommand(query, conn)
                DR = cmd.ExecuteReader
                DR.Read()
                If DR.HasRows Then
                    quantity = DR(0)
                End If
                conn.Close()
            Catch ex As MySqlException
                MsgBox("Error: " & ex.ToString())
            Finally
                conn.Close()
            End Try
            Try
                conn.Open()
                Dim cmd As New MySqlCommand()
                cmd.Connection = conn
                cmd.CommandText = "DELETE FROM outgoing_detail WHERE outgoingDetailID = " & outgoingDetailID & " AND itemID = " & itemID
                cmd.Prepare()
                cmd.ExecuteNonQuery()
                stock += quantity
                cmd.CommandText = "UPDATE raw_material SET stock = " & stock & " WHERE itemID = " & itemID
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

    Private Sub MetroButton4_Click(sender As Object, e As EventArgs) Handles MetroButton4.Click
        InLoad()
        InitLoad()
        Clear()
    End Sub

    Private Sub MetroButton5_Click(sender As Object, e As EventArgs) Handles MetroButton5.Click
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
    End Sub




    Private Sub MetroButton6_Click(sender As Object, e As EventArgs)
        If TextBox1.Text <> "" Then
            MsgBox("Mohon lakukan refresh terlebih dahulu!")
        Else
            Dim outgoingID As Integer = ComboBox3.Text
            Dim item As String = ComboBox2.Text
            Dim quantity As Integer = TextBox1.Text
            Dim stock As Integer
            Dim itemID As Integer
            Try
                conn.Open()
                Dim cmd As MySqlCommand
                Dim DR As MySqlDataReader
                Dim query As String = "SELECT itemID FROM raw_material WHERE itemName = '" & item & "'"
                cmd = New MySqlCommand(query, conn)
                DR = cmd.ExecuteReader()
                DR.Read()
                If DR.HasRows Then
                    itemID = DR(0)
                End If
            Catch ex As MySqlException
                MsgBox("Error: " & ex.ToString())
            Finally
                conn.Close()
            End Try
            Try
                conn.Open()
                Dim cmd As MySqlCommand
                Dim DR As MySqlDataReader
                Dim query As String = "SELECT stock FROM raw_material WHERE itemID = " & itemID
                cmd = New MySqlCommand(query, conn)
                DR = cmd.ExecuteReader()
                DR.Read()
                If DR.HasRows Then
                    stock = DR(0)
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
                cmd.CommandText = "INSERT INTO outgoing_detail(outgoingID, itemID, quantity) VALUES (" & outgoingID & ", " & itemID & ", " & quantity & ")"
                cmd.Prepare()
                cmd.ExecuteNonQuery()
                stock -= quantity
                cmd.CommandText = "UPDATE raw_material SET stock = " & stock & " WHERE itemID = " & itemID
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
    Private Sub UpdateBtn2_Click_1(sender As Object, e As EventArgs)

    End Sub
    Private Sub InsertBtn2_Click_1(sender As Object, e As EventArgs)

    End Sub

    Private Sub MetroButton7_Click(sender As Object, e As EventArgs)
        If TextBox1.Text = "" Then
            MsgBox("Pilih salah satu data transaksi terlebih dahulu!")
        Else
            Dim incomingDetailID As Integer = TextBox1.Text
            Dim incomingID As Integer = ComboBox3.Text
            Dim item As String = ComboBox2.Text
            Dim quantity As Integer = TextBox1.Text

            Dim itemID As Integer
            Dim initialQuantity As Integer
            Dim initialItem As Integer
            Dim initialStock As Integer
            Dim stock As Integer

            Try
                'Ambil itemID di textbox, terus ambil stoknya
                conn.Open()
                Dim cmd As MySqlCommand
                Dim DR As MySqlDataReader
                Dim query As String = "SELECT itemID, stock FROM raw_material WHERE itemName = '" & item & "'"
                cmd = New MySqlCommand(query, conn)
                DR = cmd.ExecuteReader()
                DR.Read()
                If DR.HasRows Then
                    'Dapet itemID sama stok yang mau di update
                    itemID = DR(0)
                    stock = DR(1)
                End If
            Catch ex As MySqlException
                MsgBox("Error: " & ex.ToString())
            Finally
                conn.Close()
            End Try

            Try
                'Ambil data transaksi sebelumnya
                conn.Open()
                Dim cmd As MySqlCommand
                Dim DR As MySqlDataReader
                Dim query As String = "SELECT quantity, itemID FROM incoming_detail WHERE incomingDetailID = " & incomingDetailID
                cmd = New MySqlCommand(query, conn)
                DR = cmd.ExecuteReader()
                DR.Read()
                If DR.HasRows Then
                    'Dapet itemID sama jumlah quantity transaksi lama
                    initialQuantity = DR(0)
                    initialItem = DR(1)
                End If
            Catch ex As MySqlException
                MsgBox("Error: " & ex.ToString())
            Finally
                conn.Close()
            End Try

            Try
                'Ambil data stok lama dari transaksi awal
                conn.Open()
                Dim cmd As MySqlCommand
                Dim DR As MySqlDataReader
                Dim query As String = "SELECT stock FROM raw_material WHERE itemID = " & initialItem
                cmd = New MySqlCommand(query, conn)
                DR = cmd.ExecuteReader()
                DR.Read()
                If DR.HasRows Then
                    'Dapet stok dari transaksi lama
                    initialStock = DR(0)
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
                cmd.CommandText = "UPDATE incoming_detail SET itemID = " & itemID & ", quantity = " & quantity & " WHERE incomingDetailID = " & incomingDetailID
                cmd.Prepare()
                cmd.ExecuteNonQuery()
                If initialItem <> itemID Then
                    Dim jumlahlama As Integer = initialStock - initialQuantity
                    cmd.CommandText = "UPDATE raw_material SET stock = " & jumlahlama & " WHERE itemID = " & initialItem
                    cmd.Prepare()
                    cmd.ExecuteNonQuery()
                    Dim jumlahbaru2 As Integer = stock + quantity
                    cmd.CommandText = "UPDATE raw_material SET stock = " & jumlahbaru2 & " WHERE itemID = " & itemID
                    cmd.Prepare()
                    cmd.ExecuteNonQuery()
                Else
                    Dim jumlahbaru As Integer = stock - initialQuantity + quantity
                    cmd.CommandText = "UPDATE raw_material SET stock = " & jumlahbaru & " WHERE itemID = " & itemID
                    cmd.Prepare()
                    cmd.ExecuteNonQuery()
                End If
                conn.Close()
                MsgBox("Update berhasil!")
            Catch ex As MySqlException
                MsgBox("Update gagal!")
                conn.Close()
            End Try
        End If
    End Sub

    Private Sub MetroButton8_Click(sender As Object, e As EventArgs)
        If TextBox1.Text = "" Then
            MsgBox("Pilih salah satu transaksi terlebih dahulu!")
        Else
            Dim outgoingDetailID As Integer = TextBox1.Text
            Dim item As String = ComboBox2.Text
            Dim quantity As Integer
            Dim itemID As Integer
            Dim stock As Integer
            Try
                conn.Open()
                Dim cmd As MySqlCommand
                Dim DR As MySqlDataReader
                Dim query As String = "SELECT itemID, stock FROM raw_material WHERE itemName = '" & item & "'"
                cmd = New MySqlCommand(query, conn)
                DR = cmd.ExecuteReader()
                DR.Read()
                If DR.HasRows Then
                    itemID = DR(0)
                    stock = DR(1)
                End If
            Catch ex As MySqlException
                MsgBox("Error: " & ex.ToString())
            Finally
                conn.Close()
            End Try
            Try
                conn.Open()
                Dim cmd As MySqlCommand
                Dim DR As MySqlDataReader
                Dim query As String = "SELECT quantity FROM outgoing_detail WHERE outgoingDetailID = " & outgoingDetailID
                cmd = New MySqlCommand(query, conn)
                DR = cmd.ExecuteReader
                DR.Read()
                If DR.HasRows Then
                    quantity = DR(0)
                End If
                conn.Close()
            Catch ex As MySqlException
                MsgBox("Error: " & ex.ToString())
            Finally
                conn.Close()
            End Try
            Try
                conn.Open()
                Dim cmd As New MySqlCommand()
                cmd.Connection = conn
                cmd.CommandText = "DELETE FROM outgoing_detail WHERE outgoingDetailID = " & outgoingDetailID & " AND itemID = " & itemID
                cmd.Prepare()
                cmd.ExecuteNonQuery()
                stock += quantity
                cmd.CommandText = "UPDATE raw_material SET stock = " & stock & " WHERE itemID = " & itemID
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

    Private Sub TextBox6_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub MetroButton6_Click_1(sender As Object, e As EventArgs)
        If TextBox1.Text <> "" Then
            MsgBox("Mohon lakukan refresh terlebih dahulu!")
        Else
            Dim outgoingID As Integer = ComboBox3.Text
            Dim item As String = ComboBox2.Text
            Dim quantity As Integer = TextBox1.Text
            Dim stock As Integer
            Dim itemID As Integer
            Try
                conn.Open()
                Dim cmd As MySqlCommand
                Dim DR As MySqlDataReader
                Dim query As String = "SELECT itemID FROM raw_material WHERE itemName = '" & item & "'"
                cmd = New MySqlCommand(query, conn)
                DR = cmd.ExecuteReader()
                DR.Read()
                If DR.HasRows Then
                    itemID = DR(0)
                End If
            Catch ex As MySqlException
                MsgBox("Error: " & ex.ToString())
            Finally
                conn.Close()
            End Try
            Try
                conn.Open()
                Dim cmd As MySqlCommand
                Dim DR As MySqlDataReader
                Dim query As String = "SELECT stock FROM raw_material WHERE itemID = " & itemID
                cmd = New MySqlCommand(query, conn)
                DR = cmd.ExecuteReader()
                DR.Read()
                If DR.HasRows Then
                    stock = DR(0)
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
                cmd.CommandText = "INSERT INTO outgoing_detail(outgoingID, itemID, quantity) VALUES (" & outgoingID & ", " & itemID & ", " & quantity & ")"
                cmd.Prepare()
                cmd.ExecuteNonQuery()
                stock -= quantity
                cmd.CommandText = "UPDATE raw_material SET stock = " & stock & " WHERE itemID = " & itemID
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

    Private Sub MetroButton7_Click_1(sender As Object, e As EventArgs)
        If TextBox1.Text = "" Then
            MsgBox("Pilih salah satu transaksi terlebih dahulu!")
        Else
            Dim outgoingDetailID As Integer = TextBox1.Text
            Dim outgoingID As Integer = ComboBox3.Text
            Dim item As String = TextBox1.Text
            Dim quantity As Integer = TextBox1.Text

            Dim itemID As Integer
            Dim initialQuantity As Integer
            Dim initialItem As Integer
            Dim initialStock As Integer
            Dim stock As Integer

            Try
                'Ambil itemID di textbox, terus ambil stoknya
                conn.Open()
                Dim cmd As MySqlCommand
                Dim DR As MySqlDataReader
                Dim query As String = "SELECT itemID, stock FROM raw_material WHERE itemName = '" & item & "'"
                cmd = New MySqlCommand(query, conn)
                DR = cmd.ExecuteReader()
                DR.Read()
                If DR.HasRows Then
                    'Dapet itemID sama stok yang mau di update
                    itemID = DR(0)
                    stock = DR(1)
                End If
            Catch ex As MySqlException
                MsgBox("Error: " & ex.ToString())
            Finally
                conn.Close()
            End Try

            Try
                'Ambil data transaksi sebelumnya
                conn.Open()
                Dim cmd As MySqlCommand
                Dim DR As MySqlDataReader
                Dim query As String = "SELECT quantity, itemID FROM outgoing_detail WHERE outgoingDetailID = " & outgoingDetailID
                cmd = New MySqlCommand(query, conn)
                DR = cmd.ExecuteReader()
                DR.Read()
                If DR.HasRows Then
                    'Dapet itemID sama jumlah quantity transaksi lama
                    initialQuantity = DR(0)
                    initialItem = DR(1)
                End If
            Catch ex As MySqlException
                MsgBox("Error: " & ex.ToString())
            Finally
                conn.Close()
            End Try

            Try
                'Ambil data stok lama dari transaksi awal
                conn.Open()
                Dim cmd As MySqlCommand
                Dim DR As MySqlDataReader
                Dim query As String = "SELECT stock FROM raw_material WHERE itemID = " & initialItem
                cmd = New MySqlCommand(query, conn)
                DR = cmd.ExecuteReader()
                DR.Read()
                If DR.HasRows Then
                    'Dapet stok dari transaksi lama
                    initialStock = DR(0)
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
                cmd.CommandText = "UPDATE outgoing_detail SET itemID = " & itemID & ", quantity = " & quantity & " WHERE outgoingDetailID = " & outgoingDetailID
                cmd.Prepare()
                cmd.ExecuteNonQuery()
                If initialItem <> itemID Then
                    Dim jumlahlama As Integer = initialStock + initialQuantity
                    cmd.CommandText = "UPDATE raw_material SET stock = " & jumlahlama & " WHERE itemID = " & initialItem
                    cmd.Prepare()
                    cmd.ExecuteNonQuery()
                    Dim jumlahbaru2 As Integer = stock - quantity
                    cmd.CommandText = "UPDATE raw_material SET stock = " & jumlahbaru2 & " WHERE itemID = " & itemID
                    cmd.Prepare()
                    cmd.ExecuteNonQuery()
                Else
                    Dim jumlahbaru As Integer = stock + initialQuantity - quantity
                    cmd.CommandText = "UPDATE raw_material SET stock = " & jumlahbaru & " WHERE itemID = " & itemID
                    cmd.Prepare()
                    cmd.ExecuteNonQuery()
                End If
                conn.Close()
                MsgBox("Update berhasil!")
            Catch ex As MySqlException
                MsgBox("Update gagal!")
                conn.Close()
            End Try
        End If
    End Sub



    Private Sub InsertBtn2_Click_2(sender As Object, e As EventArgs)

    End Sub

    Private Sub UpdateBtn2_Click_2(sender As Object, e As EventArgs)

    End Sub

    Private Sub MetroButton6_Click_2(sender As Object, e As EventArgs) Handles MetroButton6.Click
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
            MsgBox("Keyword tidak ditemukan")
        End If
        found = False
        TextBox5.Clear()
    End Sub

    Private Sub MetroButton7_Click_2(sender As Object, e As EventArgs) Handles MetroButton7.Click
        InitLoad()
        InLoad()
        Clear()
    End Sub

    Private Sub MetroButton8_Click_1(sender As Object, e As EventArgs) Handles MetroButton8.Click
        If TextBox1.Text <> "" Then
            MsgBox("Mohon lakukan refresh terlebih dahulu!")
        Else
            DateTimePicker1.Format = DateTimePickerFormat.Custom
            DateTimePicker1.CustomFormat = "yyyy-MM-dd"
            Dim outgoingDate As String = DateTimePicker1.Text
            Try
                conn.Open()
                Dim cmd As New MySqlCommand()
                cmd.Connection = conn
                cmd.CommandText = "INSERT INTO outgoing_item (outgoingDate) VALUES ('" & outgoingDate & "')"
                cmd.Prepare()
                cmd.ExecuteNonQuery()
                MsgBox("Insert berhasil!")
                Clear()
                conn.Close()
            Catch ex As MySqlException
                MsgBox("Insert gagal!")
                conn.Close()
            End Try
            DateTimePicker1.Format = DateTimePickerFormat.Long
        End If
    End Sub

    Private Sub MetroButton9_Click(sender As Object, e As EventArgs) Handles MetroButton9.Click
        If TextBox1.Text = "" Then
            MsgBox("Pilih salah satu data transaksi terlebih dahulu!")
        Else
            Dim outgoingID As Integer = TextBox1.Text
            Dim itemID As Integer
            Dim quantity As Integer
            Try
                Dim cmd As MySqlCommand
                Dim DR As MySqlDataReader
                conn.Open()
                Dim query As String = "SELECT itemID, quantity FROM outgoing_detail WHERE outgoingID = " & outgoingID
                cmd = New MySqlCommand(query, conn)
                DR = cmd.ExecuteReader()
                If (DR.HasRows) Then
                    While (DR.Read())
                        Dim conn2 As New MySqlConnection(constring)
                        conn2.Open()
                        Dim cmd2 As New MySqlCommand()
                        cmd2.Connection = conn2
                        itemID = DR(0)
                        quantity = DR(1)
                        cmd2.CommandText = "UPDATE raw_material SET stock = stock + " & quantity & " WHERE itemID = " & itemID
                        cmd2.Prepare()
                        cmd2.ExecuteNonQuery()
                        conn2.Close()
                    End While
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
                cmd.CommandText = "DELETE FROM outgoing_detail WHERE outgoingID = " & outgoingID
                cmd.Prepare()
                cmd.ExecuteNonQuery()
                cmd.CommandText = "DELETE FROM outgoing_item WHERE outgoingID = " & outgoingID
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

    Private Sub MetroButton10_Click(sender As Object, e As EventArgs) Handles MetroButton10.Click
        If TextBox1.Text = "" Then
            MsgBox("Pilih salah satu data transaksi terlebih dahulu!")
        Else
            Dim outgoingID As Integer = TextBox1.Text
            Dim itemID As Integer
            Dim quantity As Integer
            Try
                Dim cmd As MySqlCommand
                Dim DR As MySqlDataReader
                conn.Open()
                Dim query As String = "SELECT itemID, quantity FROM outgoing_detail WHERE outgoingID = " & outgoingID
                cmd = New MySqlCommand(query, conn)
                DR = cmd.ExecuteReader()
                If (DR.HasRows) Then
                    While (DR.Read())
                        Dim conn2 As New MySqlConnection(constring)
                        conn2.Open()
                        Dim cmd2 As New MySqlCommand()
                        cmd2.Connection = conn2
                        itemID = DR(0)
                        quantity = DR(1)
                        cmd2.CommandText = "UPDATE raw_material SET stock = stock + " & quantity & " WHERE itemID = " & itemID
                        cmd2.Prepare()
                        cmd2.ExecuteNonQuery()
                        conn2.Close()
                    End While
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
                cmd.CommandText = "DELETE FROM outgoing_detail WHERE outgoingID = " & outgoingID
                cmd.Prepare()
                cmd.ExecuteNonQuery()
                cmd.CommandText = "DELETE FROM outgoing_item WHERE outgoingID = " & outgoingID
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
End Class
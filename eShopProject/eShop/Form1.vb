Imports System.Data.OleDb


Public Class Form1
    Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;data source=C:\\Users\\PawanKumar\\Desktop\\eShopProject\\eShop\\eshop.mdb"
    Dim ds As New DataSet()




    Private Sub Label12_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Label12.Click

    End Sub


    Private Sub Button3_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button3.Click
        Dim sql As String = "SELECT * FROM tblPrd"
        Dim connection As New OleDbConnection(connectionString)
        Dim dataadapter As New OleDbDataAdapter(sql, connection)
        connection.Open()
        dataadapter.Fill(ds, "tblPrd")
        connection.Close()
        DataGridView3.DataSource = ds
        DataGridView3.DataMember = "tblPrd"

    End Sub

    Private Sub Button4_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button4.Click
        Dim sql As String = "SELECT * FROM tblOfr"
        Dim connection As New OleDbConnection(connectionString)
        Dim dataadapter As New OleDbDataAdapter(sql, connection)
        connection.Open()
        dataadapter.Fill(ds, "tblOfr")
        connection.Close()
        DataGridView1.DataSource = ds
        DataGridView1.DataMember = "tblOfr"
    End Sub

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button1.Click
        Dim insertSql As String = "INSERT INTO tblCust VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
        Using conn As New OleDbConnection(connectionString),
            cmd As New OleDbCommand(insertSql, conn)
            cmd.Parameters.Add("?", OleDbType.Integer).Value = CInt(TextBox9.Text)
            cmd.Parameters.Add("?", OleDbType.VarWChar, 20).Value = TextBox19.Text
            cmd.Parameters.Add("?", OleDbType.VarWChar, 20).Value = TextBox20.Text
            cmd.Parameters.Add("?", OleDbType.Integer).Value = CInt(TextBox21.Text)
            cmd.Parameters.Add("?", OleDbType.VarWChar, 20).Value = TextBox22.Text
            cmd.Parameters.Add("?", OleDbType.VarWChar, 20).Value = ComboBox2.Text
            cmd.Parameters.Add("?", OleDbType.Integer).Value = CInt(TextBox1.Text)
            cmd.Parameters.Add("?", OleDbType.Integer).Value = CInt(TextBox2.Text)
            cmd.Parameters.Add("?", OleDbType.VarWChar, 20).Value = DateTimePicker1.Text
            cmd.Parameters.Add("?", OleDbType.Integer).Value = CInt(TextBox4.Text)
            conn.Open()
            cmd.ExecuteNonQuery()
            MessageBox.Show("Congratulations!!! Your Order is Successful. Plz Pay = " + TextBox4.Text + " Rs.")
        End Using
    End Sub

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        ComboBox4.Items.Add("iphone12pro")
        ComboBox4.Items.Add("iphone11pro")
        ComboBox4.Items.Add("MacBook AirM1")


        Dim sql As String = "SELECT * FROM tblPrd"
        Dim connection As New OleDbConnection(connectionString)
        Dim dataadapter As New OleDbDataAdapter(sql, connection)
        connection.Open()
        dataadapter.Fill(ds, "tblPrd")
        connection.Close()
        ComboBox2.DisplayMember = "prdNme"
        ComboBox2.DataSource = ds.Tables(0)

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs)

    End Sub
    Private Sub Button5_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button5.Click
        Dim a, b, c As Integer

        b = TextBox1.Text
        c = TextBox2.Text
        a = b * c
        TextBox4.Text = a.ToString()
        ListBox1.Items.Add("Customer ID  = " + TextBox9.Text)
        ListBox1.Items.Add("Customer Name  = " + TextBox19.Text)
        ListBox1.Items.Add("Customer Phone No = " + TextBox20.Text)
        ListBox1.Items.Add("Customer Address = " + TextBox22.Text)
        ListBox1.Items.Add("Customer Product Name = " + ComboBox2.Text)
        ListBox1.Items.Add("Customer Product Price = " + TextBox1.Text)
        ListBox1.Items.Add("Customer Product Qty = " + TextBox2.Text)
        ListBox1.Items.Add("Customer Product BDate = " + DateTimePicker1.Text)
        ListBox1.Items.Add("Total Bill  = " + a.ToString())

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ComboBox2.SelectedIndexChanged


    End Sub

    Private Sub Button2_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button2.Click
        Dim sql As String = "SELECT * FROM tblCust WHERE custID=" + TextBox5.Text
        Dim connection As New OleDbConnection(connectionString)
        Dim dataadapter As New OleDbDataAdapter(sql, connection)
        connection.Open()
        dataadapter.Fill(ds, "tblCust")
        connection.Close()
        DataGridView2.DataSource = ds
        DataGridView2.DataMember = "tblCust"
    End Sub

    Private Sub Label25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label25.Click

    End Sub

    Private Sub TabPage4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage4.Click

    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim insertSql As String = "INSERT INTO tblInv VALUES (?, ?, ?, ?, ?, ?)"
        Using conn As New OleDbConnection(connectionString),
            cmd As New OleDbCommand(insertSql, conn)
            cmd.Parameters.Add("?", OleDbType.Integer).Value = CInt(TextBox10.Text)
            cmd.Parameters.Add("?", OleDbType.VarWChar, 20).Value = TextBox3.Text()
            cmd.Parameters.Add("?", OleDbType.Integer).Value = CInt(TextBox11.Text)
            cmd.Parameters.Add("?", OleDbType.VarWChar, 20).Value = DateTimePicker2.Text
            cmd.Parameters.Add("?", OleDbType.VarWChar, 20).Value = DateTimePicker3.Text
            cmd.Parameters.Add("?", OleDbType.Integer).Value = CInt(TextBox12.Text)
            conn.Open()
            cmd.ExecuteNonQuery()
            MessageBox.Show("Congratulations!!! Record Created")
        End Using
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        Dim insertSql As String = "INSERT INTO tblSup VALUES (?, ?, ?, ?, ?, ?, ?, ?)"
        Using conn As New OleDbConnection(connectionString),
            cmd As New OleDbCommand(insertSql, conn)
            cmd.Parameters.Add("?", OleDbType.Integer).Value = CInt(TextBox13.Text)
            cmd.Parameters.Add("?", OleDbType.VarWChar, 20).Value = TextBox14.Text
            cmd.Parameters.Add("?", OleDbType.VarWChar, 20).Value = TextBox15.Text
            cmd.Parameters.Add("?", OleDbType.VarWChar, 20).Value = TextBox16.Text
            cmd.Parameters.Add("?", OleDbType.VarWChar, 20).Value = ComboBox4.Text
            cmd.Parameters.Add("?", OleDbType.Integer).Value = CInt(TextBox17.Text)
            cmd.Parameters.Add("?", OleDbType.Integer).Value = CInt(TextBox18.Text)
            cmd.Parameters.Add("?", OleDbType.VarWChar, 20).Value = DateTimePicker5.Text
            conn.Open()
            cmd.ExecuteNonQuery()
            MessageBox.Show("Congratulations!!! Record Created")
        End Using
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Dim sql As String = "SELECT * FROM tblSup"
        Dim connection As New OleDbConnection(connectionString)
        Dim dataadapter As New OleDbDataAdapter(sql, connection)
        connection.Open()
        dataadapter.Fill(ds, "tblSup")
        connection.Close()
        DataGridView7.DataSource = ds
        DataGridView7.DataMember = "tblSup"
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        Dim sql As String = "DELETE FROM tblSup WHERE supID=" + TextBox6.Text
        Dim connection As New OleDbConnection(connectionString)
        Dim dataadapter As New OleDbDataAdapter(sql, connection)
        Dim ds As New DataSet()
        connection.Open()
        dataadapter.Fill(ds, "tblSup")
        connection.Close()
        MessageBox.Show("Records Deleted successfully!!!!")
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Dim sql As String = "SELECT * FROM tblInv"
        Dim connection As New OleDbConnection(connectionString)
        Dim dataadapter As New OleDbDataAdapter(sql, connection)
        connection.Open()
        dataadapter.Fill(ds, "tblInv")
        connection.Close()
        DataGridView4.DataSource = ds
        DataGridView4.DataMember = "tblInv"
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Dim sql As String = "DELETE FROM tblInv WHERE invID=" + TextBox8.Text
        Dim connection As New OleDbConnection(connectionString)
        Dim dataadapter As New OleDbDataAdapter(sql, connection)
        Dim ds As New DataSet()
        connection.Open()
        dataadapter.Fill(ds, "tblInv")
        connection.Close()
        MessageBox.Show("Records Deleted successfully!!!!")
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox4.Clear()
        TextBox9.Clear()
        TextBox19.Clear()
        TextBox20.Clear()
        TextBox21.Clear()
        TextBox22.Clear()
        ListBox1.Items.Clear()
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Dim sql As String = "SELECT custID,custNme,custPno,custAge,custAdd FROM tblCust"
        Dim connection As New OleDbConnection(connectionString)
        Dim dataadapter As New OleDbDataAdapter(sql, connection)
        connection.Open()
        dataadapter.Fill(ds, "tblCust")
        connection.Close()
        DataGridView5.DataSource = ds
        DataGridView5.DataMember = "tblCust"
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Dim sql As String = "DELETE FROM tblCust WHERE custID=" + TextBox7.Text
        Dim connection As New OleDbConnection(connectionString)
        Dim dataadapter As New OleDbDataAdapter(sql, connection)
        Dim ds As New DataSet()
        connection.Open()
        dataadapter.Fill(ds, "tblCust")
        connection.Close()
        MessageBox.Show("Records Deleted successfully!!!!")
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        Dim insertSql As String = "INSERT INTO tblOfr VALUES (?, ?, ?, ?)"
        Using conn As New OleDbConnection(connectionString),
            cmd As New OleDbCommand(insertSql, conn)
            cmd.Parameters.Add("?", OleDbType.Integer).Value = CInt(TextBox28.Text)
            cmd.Parameters.Add("?", OleDbType.VarWChar, 20).Value = TextBox30.Text()
            cmd.Parameters.Add("?", OleDbType.Integer).Value = CInt(TextBox31.Text)
            cmd.Parameters.Add("?", OleDbType.VarWChar, 20).Value = TextBox32.Text()
            conn.Open()
            cmd.ExecuteNonQuery()
            MessageBox.Show("Congratulations!!! Record Created")
        End Using
    End Sub

    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        Dim sql As String = "DELETE FROM tblOfr WHERE prdID=" + TextBox33.Text
        Dim connection As New OleDbConnection(connectionString)
        Dim dataadapter As New OleDbDataAdapter(sql, connection)
        Dim ds As New DataSet()
        connection.Open()
        dataadapter.Fill(ds, "tblOfr")
        connection.Close()
        MessageBox.Show("Records Deleted successfully!!!!")
    End Sub

    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click
        Dim insertSql As String = "INSERT INTO tblPrd VALUES (?, ?, ?, ?, ?)"
        Using conn As New OleDbConnection(connectionString),
            cmd As New OleDbCommand(insertSql, conn)
            cmd.Parameters.Add("?", OleDbType.Integer).Value = CInt(TextBox35.Text)
            cmd.Parameters.Add("?", OleDbType.VarWChar, 20).Value = TextBox36.Text()
            cmd.Parameters.Add("?", OleDbType.VarWChar, 20).Value = TextBox37.Text()
            cmd.Parameters.Add("?", OleDbType.VarWChar, 20).Value = TextBox38.Text()
            cmd.Parameters.Add("?", OleDbType.Integer).Value = CInt(TextBox39.Text)
            conn.Open()
            cmd.ExecuteNonQuery()
            MessageBox.Show("Congratulations!!! Record Created")
        End Using
    End Sub

    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        Dim sql As String = "DELETE FROM tblPrd WHERE prdID=" + TextBox34.Text
        Dim connection As New OleDbConnection(connectionString)
        Dim dataadapter As New OleDbDataAdapter(sql, connection)
        Dim ds As New DataSet()
        connection.Open()
        dataadapter.Fill(ds, "tblPrd")
        connection.Close()
        MessageBox.Show("Records Deleted successfully!!!!")
    End Sub

    Private Sub Button15_Click_1(sender As Object, e As EventArgs) Handles Button15.Click
        Dim sql As String = "SELECT custNme,custPno,custAge,custAdd FROM tblCust WHERE custID=" + TextBox9.Text
        Dim connection As New OleDbConnection(connectionString)
        connection.Open()
        Dim cmd As New OleDbCommand(sql, connection)
        Dim dr As OleDbDataReader = cmd.ExecuteReader()
        If dr.Read() Then
            TextBox19.Text = dr(0).ToString()
            TextBox20.Text = dr(1).ToString()
            TextBox21.Text = dr(2).ToString()
            TextBox22.Text = dr(3).ToString()
        Else
            MessageBox.Show("New Customer!!")
        End If
        connection.Close()
    End Sub
End Class

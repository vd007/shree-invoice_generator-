Imports System.Data.SqlClient
Imports Microsoft.Office.Interop

Public Class Form4
    Dim i As Integer
    Dim ds As New DataSet
    Dim totalrow As DataSet
    Dim MaxRows As Integer
    Dim count As Integer
    Private Sub Form4_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        i = 0

        Dim con As New OleDb.OleDbConnection
        Dim dbProvider As String
        Dim dbSource As String

        Dim da As OleDb.OleDbDataAdapter
        Dim sql As String

        Dim exec As OleDb.OleDbDataAdapter
        dbProvider = "Provider=Microsoft.ACE.OLEDB.12.0;"
        dbSource = "Data Source = C:/shree/files/user.accdb"

        con.ConnectionString = dbProvider & dbSource
        con.Open()
        sql = "select * from users"
        da = New OleDb.OleDbDataAdapter(sql, con)
        da.Fill(ds, "users")
        'count = "SELECT COUNT(*) FROM users"
        'exec = New OleDb.OleDbDataAdapter(count, con)
        'exec.Fill(totalrow, "users")
        count = ds.Tables("users").Rows.Count
        If i < count Then
            Cust_nameTextBox.Text = ds.Tables("users").Rows(i).Item(1)
            NumberTextBox.Text = ds.Tables("users").Rows(i).Item(2)
            TextBox1.Text = ds.Tables("users").Rows(i).Item(3)
            TextBox2.Text = ds.Tables("users").Rows(i).Item(4)
            i = i + 1
        End If
        If i < count Then
            TextBox3.Text = ds.Tables("users").Rows(i).Item(1)
            TextBox4.Text = ds.Tables("users").Rows(i).Item(2)
            TextBox5.Text = ds.Tables("users").Rows(i).Item(3)
            TextBox6.Text = ds.Tables("users").Rows(i).Item(4)
            i = i + 1
        End If
        If i < count Then
            TextBox7.Text = ds.Tables("users").Rows(i).Item(1)
            TextBox8.Text = ds.Tables("users").Rows(i).Item(2)
            TextBox9.Text = ds.Tables("users").Rows(i).Item(3)
            TextBox10.Text = ds.Tables("users").Rows(i).Item(4)
            i = i + 1
        End If
        If i < count Then
            TextBox11.Text = ds.Tables("users").Rows(i).Item(1)
            TextBox12.Text = ds.Tables("users").Rows(i).Item(2)
            TextBox13.Text = ds.Tables("users").Rows(i).Item(3)
            TextBox14.Text = ds.Tables("users").Rows(i).Item(4)
            i = i + 1
        End If
        If i < count Then
            TextBox15.Text = ds.Tables("users").Rows(i).Item(1)
            TextBox16.Text = ds.Tables("users").Rows(i).Item(2)
            TextBox17.Text = ds.Tables("users").Rows(i).Item(3)
            TextBox18.Text = ds.Tables("users").Rows(i).Item(4)
        End If
        
        con.Close()


    End Sub


    Private Sub Cust_nameLabel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Cust_nameTextBox_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cust_nameTextBox.TextChanged

    End Sub

    Private Sub NumberLabel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub NumberTextBox_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumberTextBox.TextChanged

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ' MaxRows = ds.Tables("users").Rows.Count
        ' If (i < MaxRows - 1) Then
        If i < count Then
            If i < count Then
                i = i + 1
                Cust_nameTextBox.Text = ds.Tables("users").Rows(i).Item(1)
                NumberTextBox.Text = ds.Tables("users").Rows(i).Item(2)
                TextBox1.Text = ds.Tables("users").Rows(i).Item(3)
                TextBox2.Text = ds.Tables("users").Rows(i).Item(4)
                i = i + 1
            End If
            If i < count Then
                TextBox3.Text = ds.Tables("users").Rows(i).Item(1)
                TextBox4.Text = ds.Tables("users").Rows(i).Item(2)
                TextBox5.Text = ds.Tables("users").Rows(i).Item(3)
                TextBox6.Text = ds.Tables("users").Rows(i).Item(4)
                i = i + 1
            End If
            If i < count Then
                TextBox7.Text = ds.Tables("users").Rows(i).Item(1)
                TextBox8.Text = ds.Tables("users").Rows(i).Item(2)
                TextBox9.Text = ds.Tables("users").Rows(i).Item(3)
                TextBox10.Text = ds.Tables("users").Rows(i).Item(4)
                i = i + 1
            End If
            If i < count Then
                TextBox11.Text = ds.Tables("users").Rows(i).Item(1)
                TextBox12.Text = ds.Tables("users").Rows(i).Item(2)
                TextBox13.Text = ds.Tables("users").Rows(i).Item(3)
                TextBox14.Text = ds.Tables("users").Rows(i).Item(4)
                i = i + 1
            End If
            If i < count Then
                TextBox15.Text = ds.Tables("users").Rows(i).Item(1)
                TextBox16.Text = ds.Tables("users").Rows(i).Item(2)
                TextBox17.Text = ds.Tables("users").Rows(i).Item(3)
                TextBox18.Text = ds.Tables("users").Rows(i).Item(4)
            End If


            'i = i + 5
            'Cust_nameTextBox.Text = ds.Tables("users").Rows(i).Item(1)
            'NumberTextBox.Text = ds.Tables("users").Rows(i).Item(2)
            'TextBox1.Text = ds.Tables("users").Rows(i).Item(3)
            'TextBox2.Text = ds.Tables("users").Rows(i).Item(4)
            'i = i + 1
            'TextBox3.Text = ds.Tables("users").Rows(i).Item(1)
            'TextBox4.Text = ds.Tables("users").Rows(i).Item(2)
            'TextBox5.Text = ds.Tables("users").Rows(i).Item(3)
            'TextBox6.Text = ds.Tables("users").Rows(i).Item(4)
            'i = i + 1
            'TextBox7.Text = ds.Tables("users").Rows(i).Item(1)
            'TextBox8.Text = ds.Tables("users").Rows(i).Item(2)
            'TextBox9.Text = ds.Tables("users").Rows(i).Item(3)
            'TextBox10.Text = ds.Tables("users").Rows(i).Item(4)
            'i = i + 1
            'TextBox11.Text = ds.Tables("users").Rows(i).Item(1)
            'TextBox12.Text = ds.Tables("users").Rows(i).Item(2)
            'TextBox13.Text = ds.Tables("users").Rows(i).Item(3)
            'TextBox14.Text = ds.Tables("users").Rows(i).Item(4)
            'i = i + 1
            'TextBox15.Text = ds.Tables("users").Rows(i).Item(1)
            'TextBox16.Text = ds.Tables("users").Rows(i).Item(2)
            'TextBox17.Text = ds.Tables("users").Rows(i).Item(3)
            'TextBox18.Text = ds.Tables("users").Rows(i).Item(4)
        Else
            MsgBox("end of records")

        End If
        


    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If (i > 0) Then
            If i > 0 Then
                i = i - 1
                Cust_nameTextBox.Text = ds.Tables("users").Rows(i).Item(1)
                NumberTextBox.Text = ds.Tables("users").Rows(i).Item(2)
                TextBox1.Text = ds.Tables("users").Rows(i).Item(3)
                TextBox2.Text = ds.Tables("users").Rows(i).Item(4)
                i = i - 1
            End If
            If i > 0 Then
                TextBox3.Text = ds.Tables("users").Rows(i).Item(1)
                TextBox4.Text = ds.Tables("users").Rows(i).Item(2)
                TextBox5.Text = ds.Tables("users").Rows(i).Item(3)
                TextBox6.Text = ds.Tables("users").Rows(i).Item(4)
                i = i - 1
            End If
            If i > 0 Then
                TextBox7.Text = ds.Tables("users").Rows(i).Item(1)
                TextBox8.Text = ds.Tables("users").Rows(i).Item(2)
                TextBox9.Text = ds.Tables("users").Rows(i).Item(3)
                TextBox10.Text = ds.Tables("users").Rows(i).Item(4)
                i = i - 1
            End If
            If i > 0 Then
                TextBox11.Text = ds.Tables("users").Rows(i).Item(1)
                TextBox12.Text = ds.Tables("users").Rows(i).Item(2)
                TextBox13.Text = ds.Tables("users").Rows(i).Item(3)
                TextBox14.Text = ds.Tables("users").Rows(i).Item(4)
                i = i - 1
            End If
            If i > 0 Then

                TextBox15.Text = ds.Tables("users").Rows(i).Item(1)
                TextBox16.Text = ds.Tables("users").Rows(i).Item(2)
                TextBox17.Text = ds.Tables("users").Rows(i).Item(3)
                TextBox18.Text = ds.Tables("users").Rows(i).Item(4)
            End If

        Else
            MsgBox("end of records")
        End If
    End Sub
End Class
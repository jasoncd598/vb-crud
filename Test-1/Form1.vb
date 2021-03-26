Imports System.Data.SqlClient
Imports Microsoft.VisualBasic
Public Class Form1
    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim ID As Integer = TextBox6.Text
        Dim NAME As String = TextBox1.Text
        Dim FIRSTNAME As String = TextBox2.Text
        Dim LASTNAME As String = TextBox3.Text
        Dim GENDER As String = ComboBox1.Text
        Dim PHONENUMBER As Integer = TextBox4.Text

        Dim query As String = "INSERT INTO member VALUES (@ID, @NAME, @FIRSTNAME, @LASTNAME, @GENDER, @PHONENUMBER)"
        Using conn As SqlConnection = New SqlConnection("Data Source=.;Initial Catalog=test-dbase;Integrated Security=True")
            Using cmd As SqlCommand = New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@ID", ID)
                cmd.Parameters.AddWithValue("@NAME", NAME)
                cmd.Parameters.AddWithValue("@FIRSTNAME", FIRSTNAME)
                cmd.Parameters.AddWithValue("@LASTNAME", LASTNAME)
                cmd.Parameters.AddWithValue("@GENDER", GENDER)
                cmd.Parameters.AddWithValue("@PHONENUMBER", PHONENUMBER)

                conn.Open()
                cmd.ExecuteNonQuery()
                conn.Close()
                TextBox1.Text = ""
                TextBox2.Text = ""
                TextBox3.Text = ""
                TextBox4.Text = ""
                TextBox6.Text = ""
                MessageBox.Show("A new record has been saved!")
                BindData()
            End Using
        End Using
    End Sub

    Public Sub BindData()
        Dim query As String = "SELECT * FROM member"
        Using conn As SqlConnection = New SqlConnection("Data Source=.;Initial Catalog=test-dbase;Integrated Security=True")
            Using cmd As SqlCommand = New SqlCommand(query, conn)
                Using da As New SqlDataAdapter()
                    da.SelectCommand = cmd
                    Using dt As New DataTable()
                        da.Fill(dt)
                        DataGridView1.DataSource = dt


                    End Using

                End Using
                End Using
        End Using
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim query As String = "SELECT * FROM member WHERE ID ='" & TextBox5.Text & "'"
        Using conn As SqlConnection = New SqlConnection("Data Source=.;Initial Catalog=test-dbase;Integrated Security=True")
            Using cmd As SqlCommand = New SqlCommand(query, conn)
                Using da As New SqlDataAdapter()
                    da.SelectCommand = cmd
                    Using dt As New DataTable()
                        da.Fill(dt)
                        If dt.Rows.Count > 0 Then
                            TextBox6.Text = dt.Rows(0)(0).ToString()
                            TextBox1.Text = dt.Rows(0)(1).ToString()
                            TextBox2.Text = dt.Rows(0)(2).ToString()
                            TextBox3.Text = dt.Rows(0)(3).ToString()
                            TextBox4.Text = dt.Rows(0)(5).ToString()
                            ComboBox1.Text = dt.Rows(0)(4).ToString()
                        Else
                            TextBox1.Text = ""
                            TextBox2.Text = ""
                            TextBox3.Text = ""
                            TextBox4.Text = ""
                            TextBox6.Text = ""

                        End If
                    End Using
                End Using
            End Using
        End Using
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim ID As Integer = TextBox6.Text
        Dim NAME As String = TextBox1.Text
        Dim FIRSTNAME As String = TextBox2.Text
        Dim LASTNAME As String = TextBox3.Text
        Dim GENDER As String = ComboBox1.Text
        Dim PHONENUMBER As Integer = TextBox4.Text

        Dim query As String = "UPDATE member SET  U_NAME=@NAME, U_FIRST_NAME=@FIRSTNAME, U_LAST_NAME=@LASTNAME, U_GENDER=@GENDER, U_PHONE_NUMBER=@PHONENUMBER WHERE ID=@ID"
        Using conn As SqlConnection = New SqlConnection("Data Source=.;Initial Catalog=test-dbase;Integrated Security=True")
            Using cmd As SqlCommand = New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@ID", ID)
                cmd.Parameters.AddWithValue("@NAME", NAME)
                cmd.Parameters.AddWithValue("@FIRSTNAME", FIRSTNAME)
                cmd.Parameters.AddWithValue("@LASTNAME", LASTNAME)
                cmd.Parameters.AddWithValue("@GENDER", GENDER)
                cmd.Parameters.AddWithValue("@PHONENUMBER", PHONENUMBER)

                conn.Open()
                cmd.ExecuteNonQuery()
                conn.Close()
                TextBox1.Text = ""
                TextBox2.Text = ""
                TextBox3.Text = ""
                TextBox4.Text = ""
                TextBox6.Text = ""
                MessageBox.Show("A record has been updated!")
                BindData()
            End Using
        End Using
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim ID As String = TextBox6.Text
        Dim query As String = "DELETE FROM member WHERE ID =@ID"
        Using conn As SqlConnection = New SqlConnection("Data Source=.;Initial Catalog=test-dbase;Integrated Security=True")
            Using cmd As SqlCommand = New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@ID", ID)

                conn.Open()
                cmd.ExecuteNonQuery()
                conn.Close()
                TextBox1.Text = ""
                TextBox2.Text = ""
                TextBox3.Text = ""
                TextBox4.Text = ""
                TextBox6.Text = ""
                MessageBox.Show("A record has been deleted!")
                BindData()
            End Using
        End Using
    End Sub
    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        Dim ch As Char = e.KeyChar
        If Not Char.IsDigit(ch) AndAlso Asc(ch) <> 8 Then
            e.Handled = True
        End If
    End Sub


    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        RegexString(e)
    End Sub

    Public Sub RegexString(e)
        If Not (Asc(e.KeyChar) = 8) Then
            Dim allowedChars As String = "abcdefghijklmnopqrstuvwxyz"
            If Not allowedChars.Contains(e.KeyChar.ToString.ToLower) Then
                e.KeyChar = ChrW(0)
                e.Handled = True
            End If
        End If
    End Sub
    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        RegexString(e)
    End Sub

    Private Sub TextBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress
        RegexString(e)
    End Sub
End Class

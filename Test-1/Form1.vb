Imports System.Data.SqlClient

Public Class Form1

    Dim DataSource As String = "Data Source=.;Initial Catalog=test-dbase;Integrated Security=True"

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Getting values of the field in the form'
        Dim ID As Integer = Val(TextBox6.Text)
        Dim NAME As String = TextBox1.Text
        Dim FIRSTNAME As String = TextBox2.Text
        Dim LASTNAME As String = TextBox3.Text
        Dim GENDER As String = ComboBox1.Text
        Dim PHONENUMBER As Integer = Val(TextBox4.Text)

        'Validate if fields are empty, if does saving process is interrupted.'
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox6.Text = "" Or String.IsNullOrEmpty(ComboBox1.Text) Then
            MessageBox.Show("Please fill the empty field.")
        Else
            Dim query As String = "INSERT INTO member VALUES (@ID, @NAME, @FIRSTNAME, @LASTNAME, @GENDER, @PHONENUMBER)"
            Using conn As SqlConnection = New SqlConnection(DataSource)
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
                    ClearField()
                    MessageBox.Show("A new record has been saved!")
                    BindData()
                End Using
            End Using
        End If
    End Sub

    Public Sub BindData()
        Dim query As String = "SELECT * FROM member"
        Using conn As SqlConnection = New SqlConnection(DataSource)
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
        Using conn As SqlConnection = New SqlConnection(DataSource)
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

                            ClearField()
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

        Dim query As String = "UPDATE member SET U_NAME=@NAME, U_FIRST_NAME=@FIRSTNAME, U_LAST_NAME=@LASTNAME, U_GENDER=@GENDER, U_PHONE_NUMBER=@PHONENUMBER WHERE ID=@ID"
        Using conn As SqlConnection = New SqlConnection(DataSource)
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

                ClearField()

                MessageBox.Show("A record has been updated!")

                BindData()
            End Using
        End Using
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim ID As String = TextBox6.Text
        Dim query As String = "DELETE FROM member WHERE ID =@ID"
        Using conn As SqlConnection = New SqlConnection(DataSource)
            Using cmd As SqlCommand = New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@ID", ID)

                conn.Open()
                cmd.ExecuteNonQuery()
                conn.Close()

                ClearField()

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

    Public Sub RegexString(e, sender)
        If (Not Char.IsLetter(e.KeyChar) AndAlso e.KeyChar <> " "c) Then e.Handled = True
        If (DirectCast(sender, TextBox).Text.Length = 0) Then e.KeyChar = Char.ToUpper(e.KeyChar)
    End Sub

    Public Sub ClearField()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox6.Text = ""
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        RegexString(e, sender)
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        RegexString(e, sender)
    End Sub

    Private Sub TextBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress
        RegexString(e, sender)
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        BindData()
    End Sub
End Class

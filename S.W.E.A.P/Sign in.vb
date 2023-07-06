﻿Imports MySql.Data.MySqlClient

Public Class Form2
    ' Connection object
    Public Shared conn As MySqlConnection

    Private rid As MySqlDataReader

    ' Server change form
    Private serverChange As serverChange

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Create the initial connection object
        conn = New MySqlConnection("server=172.30.206.180;port=3306;username=dswd;password=sweap123;database=sweap")
    End Sub

    Private Sub btnChangeServer_Click(sender As Object, e As EventArgs) Handles btnChangeServer.Click
        ' Create a new instance of the server change form
        serverChange = New serverChange()

        ' Show the server change form as a dialog
        If serverChange.ShowDialog() = DialogResult.OK Then
            ' Retrieve the new server details from the form
            Dim newServer As String = serverChange.txtServer.Text
            Dim newPort As Integer = Integer.Parse(serverChange.txtPort.Text)

            ' Update the connection string with the new server details
            conn.ConnectionString = $"server={newServer};port={newPort};username=dswd;password=sweap123;database=sweap"

            ' Show a message indicating successful server change
            MessageBox.Show("Server changed successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

        ' Dispose the server change form
        serverChange.Dispose()
    End Sub

    Public Shared log_id As Integer

    'Perform log-in by using enter.
    Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click
        PerformLogin()
    End Sub

    Private Sub txtPassword_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPassword.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            PerformLogin()
            e.Handled = True
        End If
    End Sub

    Private Sub PerformLogin()
        Dim DateAndTime As String = DateTime.Now.ToString()

        If (txtUsername.Text = "") Then
            MessageBox.Show("Username can't be blank.", "Invalid", MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf (txtPassword.Text = "") Then
            MessageBox.Show("Password can't be blank.", "Invalid", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            Dim status As Integer = 3
            Try
                conn.Open()
                Dim cmd As New MySqlCommand("SELECT * FROM users WHERE username=@NAME AND password=@PASS", conn)
                cmd.Parameters.AddWithValue("@NAME", txtUsername.Text)
                cmd.Parameters.AddWithValue("@PASS", txtPassword.Text)
                rid = cmd.ExecuteReader
                While rid.Read
                    status = rid.GetInt32("is_admin")
                    log_id = rid.GetInt32("id")
                End While
                If (status = 1) Then
                    txtUsername.Clear()
                    txtPassword.Clear()
                    admindash.Show()
                    Me.Hide()
                ElseIf (status = 0) Then
                    txtUsername.Clear()
                    txtPassword.Clear()
                    Me.Hide()
                    user_dashboard.Show()
                Else
                    MessageBox.Show("Invalid Username or Password", "Invalid", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
                rid.Close() ' Close the data reader

                ' Update the last_logout column for the logged-in user
                Dim cmdd As New MySqlCommand("UPDATE users SET last_logout=@now WHERE id = @ID", conn)
                cmdd.Parameters.AddWithValue("@ID", log_id)
                cmdd.Parameters.AddWithValue("@now", DateAndTime)
                cmdd.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message)
            Finally
                conn.Close()
            End Try
        End If

    End Sub



    Private Sub checkShowPw_CheckedChanged(sender As Object, e As EventArgs) Handles checkShowPw.CheckedChanged
        If checkShowPw.Checked = False Then
            txtPassword.PasswordChar = "*"
        Else
            txtPassword.PasswordChar = ""
        End If
    End Sub

    'REGISTER FORM
    Private Sub lblRegister_Click(sender As Object, e As EventArgs) Handles lblRegister.Click
        Me.Hide()
        signups.Show()
    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtPassword.PasswordChar = "*"
    End Sub

    Private Sub Guna2PictureBox1_Click(sender As Object, e As EventArgs) Handles Guna2PictureBox1.Click
        Me.Close()
    End Sub

    'ANOTHER FORM TO CHANGE PASS
    Private Sub lblForgot_Click(sender As Object, e As EventArgs) Handles lblForgot.Click
        ForgotPass.Show()
        Me.Hide()
    End Sub
End Class
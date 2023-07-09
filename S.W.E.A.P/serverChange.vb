Imports MySql.Data.MySqlClient

Public Class serverChange
    Private Sub Guna2Button1_Click(sender As Object, e As EventArgs) Handles Guna2Button1.Click
        Form2.query = "server=" & txtIp.Text & ";port=" & txtPort.Text & ";username=" & txtUsername.Text & ";password=" & txtPassword.Text & ";database=" & txtDatabase.Text
        Form2.conn.ConnectionString = Form2.query
        Try
            Form2.conn.Open()
            Dim location As String = My.Application.Info.DirectoryPath
            'Dim indext As Integer = locateProject.IndexOf("bin\Debug\net6.0-windows")
            'Dim location As String = locateProject.Substring(0, indext)
            Dim filepath As String = location & "\Others\lastPort.txt"
            'clears the text file of the last log ports addresses
            Using writer As New System.IO.StreamWriter(filepath, False)
                writer.Write(String.Empty)
            End Using
            'records the last log ports addresses
            Using writer As New System.IO.StreamWriter(filepath)
                writer.WriteLine(txtIp.Text)
                writer.WriteLine(txtPort.Text)
                writer.WriteLine(txtUsername.Text)
                writer.WriteLine(txtPassword.Text)
                writer.WriteLine(txtDatabase.Text)
            End Using
            MessageBox.Show("Connected Successfully!", "Response", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Form2.Show()
            Me.Close()
        Catch ex As Exception
            MessageBox.Show("Connection Failed!", "Response", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Form2.conn.Close()
        End Try

    End Sub

    Private Sub serverChange_Load(sender As Object, e As EventArgs) Handles MyBase.Load '-----AUTOLOAD
        txtPort.Text = "3306"
        txtDatabase.Text = "sweap"
    End Sub

    Private Sub Guna2Button2_Click(sender As Object, e As EventArgs) Handles Guna2Button2.Click
        Form2.Show()
        Me.Close()
    End Sub
End Class
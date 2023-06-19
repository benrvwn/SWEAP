﻿Imports DocumentFormat.OpenXml.Vml.Spreadsheet
Imports MySql.Data.MySqlClient
Imports OfficeOpenXml.Table.PivotTable
Imports System.IO
Imports System.Text.RegularExpressions
Public Class signups
    Dim conn As New MySqlConnection("server=172.30.207.132;port=3306;username=sweapp;password=druguser;database=sweap")
    Dim rid As MySqlDataReader
    Dim error_msg(0) As String
    Dim random As Integer = 0
    Dim i As Integer = 0
    Dim message As String
    Dim sourceFilePath As String
    Dim getExtension As String

    Private Sub signups_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TabPage2.Enabled = False
        TabPage3.Enabled = False
        TabPage3.Enabled = False
        TabPage4.Enabled = False
    End Sub
    Public Sub valid_blank(field, name, fieldname)
        If field = "" Then
            fieldname.BorderColor = Color.FromArgb(255, 0, 0)
            fieldname.BorderThickness = 1.5
            error_msg(random) = name & " can't be blank." & vbNewLine
            random = random + 1
            ReDim Preserve error_msg(random)
        End If
    End Sub
    Public Sub invalid_reset(fieldname)
        fieldname.bordercolor = Color.Gray
        fieldname.borderthickness = 1
    End Sub

    Public Sub reset_all()
        invalid_reset(txtbxUser)
        invalid_reset(txtbxPass)
        invalid_reset(txtFname)
        invalid_reset(txtMname)
        invalid_reset(txtLname)
        invalid_reset(comboPos)
        invalid_reset(txtAddress)
        invalid_reset(txtContact)
        invalid_reset(txtEmail)
        invalid_reset(txtEduc)
        invalid_reset(comboOffice)
        invalid_reset(comboEmployStat)
        invalid_reset(comboCommit)
        invalid_reset(txtbxBF1)
        invalid_reset(txtbxBR1)
        invalid_reset(txtBA1)
    End Sub
    Public Sub add_benefi(hook, bname, brel, bage)
        Try
            conn.Open()
            Dim cmd As New MySqlCommand("insert into beneficiaries(user_id, full_name, relationship, age)values((select id from users where username=@HOOK1), @BNAME, @BREL, @BAGE);", conn)
            cmd.Parameters.AddWithValue("@HOOK1", hook)
            cmd.Parameters.AddWithValue("@BNAME", bname)
            cmd.Parameters.AddWithValue("@BREL", brel)
            cmd.Parameters.AddWithValue("@BAGE", bage)
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("beneficiary function doesn't work.")
        Finally
            conn.Close()
        End Try
    End Sub
    Private Sub Guna2Button1_Click(sender As Object, e As EventArgs) Handles Guna2Button1.Click '-----NEXT BUTTON 1------
        valid_blank(txtFname.Text, "First name", txtFname)
        valid_blank(txtMname.Text, "Middle name", txtMname)
        valid_blank(txtLname.Text, "Last name", txtLname)
        valid_blank(pickSex.SelectedItem, "Sex", pickSex)
        valid_blank(txtAddress.Text, "Address", txtAddress)
        valid_blank(txtContact.Text, "Contact", txtContact)
        valid_blank(txtEmail.Text, "Email", txtEmail)
        valid_blank(txtEduc.Text, "Educational Attainment", txtEduc)

        While i < error_msg.Length
            message = message & error_msg(i)
            i = i + 1
        End While

        If message = "" Then
            '----------------------------NEXT-FORM-------2------------------------------------------'
            Guna2TabControl1.SelectedTab = TabPage2
            TabPage2.Enabled = True
        Else
            MsgBox(message)
            i = 0
            message = ""
            Array.Clear(error_msg, 0, error_msg.Length)
        End If
    End Sub


    Private Sub Guna2Button2_Click(sender As Object, e As EventArgs) Handles Guna2Button2.Click '-------NEXT BUTTON 2------------
        valid_blank(comboPos.SelectedItem, "Position", comboPos)
        valid_blank(comboOffice.SelectedItem, "Office", comboOffice)
        valid_blank(comboEmployStat.SelectedItem, "Employment type", comboEmployStat)
        valid_blank(comboCommit.SelectedItem, "Committee", comboCommit)
        While i < error_msg.Length
            message = message & error_msg(i)
            i = i + 1
        End While

        If message = "" Then
            '----------------------------NEXT-FORM------3-------------------------------------------'
            Guna2TabControl1.SelectedTab = TabPage3
            TabPage3.Enabled = True
        Else
            MsgBox(message)
            i = 0
            message = ""
            Array.Clear(error_msg, 0, error_msg.Length)
        End If
    End Sub


    Private Sub Guna2Button3_Click(sender As Object, e As EventArgs) Handles Guna2Button3.Click  '-------------NEXT BUTTON 3-------------------------'
        valid_blank(txtbxBF1.Text, "Beneficiarie's name", txtbxBF1)
        valid_blank(txtbxBR1.Text, "Beneficiarie's relationship", txtbxBR1)
        valid_blank(txtBA1.Text, "Beneficiarie's age", txtBA1)

        While i < error_msg.Length
            message = message & error_msg(i)
            i = i + 1
        End While

        If message = "" Then
            '----------------------------NEXT-FORM-----4--------------------------------------------'
            Guna2TabControl1.SelectedTab = TabPage4
            TabPage4.Enabled = True
        Else
            MsgBox(message)
            i = 0
            message = ""
            Array.Clear(error_msg, 0, error_msg.Length)
        End If
    End Sub

    Private Sub Guna2Button4_Click(sender As Object, e As EventArgs) Handles Guna2Button4.Click '-------SUBMIT BUTTON------------
        valid_blank(txtbxUser.Text, "Beneficiarie's name", txtbxUser)
        valid_blank(txtbxPass.Text, "Beneficiarie's relationship", txtbxPass)
        'valid_blank(pBoxProfile.Image, "Profile", pBoxProfile)

        While i < error_msg.Length
            message = message & error_msg(i)
            i = i + 1
        End While
        If message = "" Then
            '----------------------------GETTING IMAGE-------------------------------------------------
            Dim locateProject As String = My.Application.Info.DirectoryPath
            Dim indext As Integer = locateProject.IndexOf("bin\Debug\net6.0-windows")
            Dim location As String = locateProject.Substring(0, indext)
            Dim opf As New OpenFileDialog

            Dim destinationPath As String = location & "\Resources\user_profile\" & txtbxUser.Text & getExtension
            File.Copy(sourceFilePath, destinationPath, 0)
            Dim imageInput As String = "\" & txtbxUser.Text & getExtension
            '------------------------------------------------------------------------------------------
            Try
                conn.Open()
                Dim cmd As New MySqlCommand("insert into users(username, password, first_name, middle_name, last_name, position, image, is_admin, created_at, updated_at)values(@UNAME, @PW, @FNAME, @MNAME, @LNAME, @POS, @IMG, 0, now(), now());
                                         insert into user_info(user_id, address, contact, email, educational, birthdate, office, employment_status, committee)values(last_insert_id(), @ADRS, @CONTACT, @EMAIL, @EDUC, @BDAY, @OFFICE, @EMTYPE, @COMM);
                                         insert into beneficiaries(user_id, full_name, relationship, age)values((select id from users where username=@HOOK), @BNAME1, @BREL1, @BAGE1);", conn)
                cmd.Parameters.AddWithValue("@UNAME", txtbxUser.Text)
                cmd.Parameters.AddWithValue("@PW", txtbxPass.Text)
                cmd.Parameters.AddWithValue("@FNAME", txtFname.Text)
                cmd.Parameters.AddWithValue("@MNAME", txtMname.Text)
                cmd.Parameters.AddWithValue("@LNAME", txtLname.Text)
                cmd.Parameters.AddWithValue("@POS", comboPos.SelectedItem)
                cmd.Parameters.AddWithValue("@IMG", imageInput)
                cmd.Parameters.AddWithValue("@ADRS", txtAddress.Text)
                cmd.Parameters.AddWithValue("@CONTACT", txtContact.Text)
                cmd.Parameters.AddWithValue("@EMAIL", txtEmail.Text)
                cmd.Parameters.AddWithValue("@EDUC", txtEduc.Text)
                cmd.Parameters.AddWithValue("@BDAY", dateBday.Value)
                cmd.Parameters.AddWithValue("@OFFICE", comboOffice.SelectedItem)
                cmd.Parameters.AddWithValue("@EMTYPE", comboEmployStat.SelectedItem)
                cmd.Parameters.AddWithValue("@COMM", comboCommit.SelectedItem)
                cmd.Parameters.AddWithValue("@BNAME1", txtbxBF1.Text)
                cmd.Parameters.AddWithValue("@BREL1", txtbxBR1.Text)
                cmd.Parameters.AddWithValue("@BAGE1", txtBA1.Text)
                cmd.Parameters.AddWithValue("@HOOK", txtbxUser.Text)
                cmd.ExecuteNonQuery()
                MsgBox("Added successfully!")
            Catch ex As Exception
                'MsgBox("Create account failed.")
            Finally
                conn.Close()
            End Try
            If txtbxBF2.Text <> "" Then
                add_benefi(txtbxUser.Text, txtbxBF2.Text, txtbxBR2.Text, txtBA2.Text)
            End If
            If txtbxBF3.Text <> "" Then
                add_benefi(txtbxUser.Text, txtbxBF3.Text, txtbxBR3.Text, txtBA3.Text)
            End If
            If txtbxBF4.Text <> "" Then
                add_benefi(txtbxUser.Text, txtbxBF4.Text, txtbxBR4.Text, txtBA4.Text)
            End If
            If txtbxBF5.Text <> "" Then
                add_benefi(txtbxUser.Text, txtbxBF5.Text, txtbxBR5.Text, txtBA5.Text)
            End If
            Form1.Visible = True
            Me.Visible = False
        Else
            MsgBox(message)
            i = 0
            message = ""
            Array.Clear(error_msg, 0, error_msg.Length)
        End If
    End Sub

    Private Sub bttnUpload_Click(sender As Object, e As EventArgs) Handles bttnUpload.Click
        Dim opf As New OpenFileDialog

        opf.Filter = "Choose Image(*.jpg; *.png; *.gif) | * .jpg; *.png; *.gif"
        If opf.ShowDialog = DialogResult.OK Then
            'imageInput = System.IO.Path.GetFullPath(opf.FileName)
            sourceFilePath = Path.GetFullPath(opf.FileName)
            pBoxProfile.BackgroundImage = Image.FromFile(sourceFilePath)
            getExtension = Path.GetExtension(opf.FileName)
        End If
    End Sub

    '' txt type number only
    Private Sub txtContact_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtContact.KeyPress
        If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtBA1_KeyPress(sender As Object, e As EventArgs) Handles txtBA1.KeyPress
        txtContact_KeyPress(sender, e)
    End Sub

    Private Sub txtBA2_KeyPress(sender As Object, e As EventArgs) Handles txtBA2.KeyPress
        txtContact_KeyPress(sender, e)
    End Sub

    Private Sub txtBA3_KeyPress(sender As Object, e As EventArgs) Handles txtBA3.KeyPress
        txtContact_KeyPress(sender, e)
    End Sub

    Private Sub txtBA4_KeyPress(sender As Object, e As EventArgs) Handles txtBA4.KeyPress
        txtContact_KeyPress(sender, e)
    End Sub

    Private Sub txtBA5_Keypress(sender As Object, e As EventArgs) Handles txtBA5.KeyPress
        txtContact_KeyPress(sender, e)
    End Sub


    ''Email validation
    Private Function IsValidGmail(email As String) As Boolean
        Dim gmailRegex As New Regex("^[a-zA-Z0-9_.+-]+@gmail\.com$", RegexOptions.IgnoreCase)
        Return gmailRegex.IsMatch(email)
    End Function

    Private Sub txtEmail_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtEmail.Validating
        Dim inputEmail As String = txtEmail.Text.Trim()
        If txtEmail.Text = "" Then
            txtEmail.Text = txtEmail.Text


        ElseIf Not IsValidGmail(inputEmail) Then
            MessageBox.Show("Invalid Email account." & vbCrLf & "Please enter a valid Gmail address.", "INFORMATION", MessageBoxButtons.OK, MessageBoxIcon.Hand)
            e.Cancel = True

        End If
    End Sub


    '' KEY PRESS LANG TO PRE
    Private Shared Sub txtFname_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtFname.KeyPress
        If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsLetter(e.KeyChar) AndAlso Not Char.IsWhiteSpace(e.KeyChar) AndAlso Not Char.IsPunctuation(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtMname_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtMname.KeyPress
        txtFname_KeyPress(sender, e)
    End Sub

    Private Sub txtLname_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtLname.KeyPress
        txtFname_KeyPress(sender, e)
    End Sub

    Private Sub txtbxBF1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtbxBF1.KeyPress
        txtFname_KeyPress(sender, e)
    End Sub

    Private Sub txtbxBF2_KeyPress(sender As Object, e As EventArgs) Handles txtbxBF2.KeyPress
        txtFname_KeyPress(sender, e)
    End Sub

    Private Sub txtbxBF3_KeyPress(sender As Object, e As EventArgs) Handles txtbxBF3.KeyPress
        txtFname_KeyPress(sender, e)
    End Sub

    Private Sub txtbxBF4_KeyPress(sender As Object, e As EventArgs) Handles txtbxBF4.KeyPress
        txtFname_KeyPress(sender, e)
    End Sub

    Private Sub txtbxBF5_KeyPress(sender As Object, e As EventArgs) Handles txtbxBF5.KeyPress
        txtFname_KeyPress(sender, e)
    End Sub

    Private Sub txtbxBR1_KeyPress(sender As Object, e As EventArgs) Handles txtbxBR1.KeyPress
        txtFname_KeyPress(sender, e)
    End Sub

    Private Sub txtbxBR2_KeyPress(sender As Object, e As EventArgs) Handles txtbxBR2.KeyPress
        txtFname_KeyPress(sender, e)
    End Sub

    Private Sub txtbxBR3_KeyPress(sender As Object, e As EventArgs) Handles txtbxBR3.KeyPress
        txtFname_KeyPress(sender, e)
    End Sub

    Private Sub txtbxBR4_KeyPress(sender As Object, e As EventArgs) Handles txtbxBR4.KeyPress
        txtFname_KeyPress(sender, e)
    End Sub

    Private Sub txtbxBR5_KeyPress(sender As Object, e As EventArgs) Handles txtbxBR5.KeyPress
        txtFname_KeyPress(sender, e)
    End Sub
End Class
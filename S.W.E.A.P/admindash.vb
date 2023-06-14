﻿Imports TheArtOfDevHtmlRenderer.Adapters.Entities
Imports MySql.Data.MySqlClient
Imports System.IO

Public Class admindash
    Dim conn As New MySqlConnection("server=172.30.205.208;port=3306;username=sweapp;password=druguser;database=sweap")
    Dim rid As MySqlDataReader
    Dim selectedId As Integer = 0
    Dim selectedBenId As Integer
    Dim currentBen As Integer

    Public Sub countBen() '------TO COUNT BENEFICIARIES OF SPECIFIC USER(FOR EDITING PURPOSES)
        Try
            conn.Open()
            Dim cmd As New MySqlCommand("select count(id) as counted from beneficiaries where user_id=@ID", conn)
            cmd.Parameters.AddWithValue("@ID", selectedId)
            rid = cmd.ExecuteReader
            While rid.Read
                currentBen = rid.GetInt32("counted")
            End While
        Catch ex As Exception
        Finally
            conn.Close()
        End Try
    End Sub

    Public Sub beneficiariesRecord() '---------------FOR BENEFICIARIES RECORD
        dgBeneficiaries.Rows.Clear()
        Try
            conn.Open()
            Dim cmd As New MySqlCommand("select * from beneficiaries where user_id = @ID", conn)
            cmd.Parameters.AddWithValue("@ID", selectedId)
            rid = cmd.ExecuteReader
            While rid.Read
                dgBeneficiaries.Rows.Add(rid.Item("id"), rid.Item("full_name"), rid.Item("relationship"), rid.Item("age"))
            End While
        Catch ex As Exception
        Finally
            conn.Close()
        End Try
    End Sub
    Public Sub viewMembers(query) '-----------------FOR EMPLOYEES TABLE
        dgMembers.Rows.Clear()
        Try
            conn.Open()
            Dim cmd As New MySqlCommand(query, conn)
            rid = cmd.ExecuteReader
            While rid.Read
                dgMembers.Rows.Add(rid.Item("id"), rid.Item("full_name"), rid.Item("office"), rid.Item("position"), rid.Item("employment_status"), rid.Item("email"))
            End While
        Catch ex As Exception
        Finally
            conn.Close()
        End Try
    End Sub

    Private Sub btnBenRemove_Click(sender As Object, e As EventArgs) Handles btnBenRemove.Click
        Dim result As DialogResult = MessageBox.Show("Are you sure you want to remove " & dgBeneficiaries.CurrentRow.Cells(1).Value.ToString() &
                                                      " from beneficiaries?", "Confirmation", MessageBoxButtons.YesNo)
        If result = DialogResult.Yes Then
            Try
                conn.Open()
                Dim cmd As New MySqlCommand("delete from beneficiaries where id=@ID", conn)
                cmd.Parameters.AddWithValue("@ID", selectedBenId)
                cmd.ExecuteNonQuery()
                MsgBox("Record has been removed.")
                beneficiariesRecord()
                countBen()
            Catch ex As Exception
                MsgBox("operation failed")
            Finally
                conn.Close()
            End Try
        End If
        If currentBen < 5 Then
            txtEditAddBen.Enabled = True
            txtEditAddBenAge.Enabled = True
            txtEditAddBenRel.Enabled = True
            btnEditAddBen.Enabled = True
        End If
        beneficiariesRecord()
    End Sub

    Private Sub btnEditAddBen_Click(sender As Object, e As EventArgs) Handles btnEditAddBen.Click
        countBen()
        If currentBen < 5 Then
            Try
                conn.Open()
                Dim cmd As New MySqlCommand("insert into beneficiaries(user_id, full_name, relationship, age) values(@ID, @FNAME, @REL, @AGE)", conn)
                cmd.Parameters.AddWithValue("@ID", selectedId)
                cmd.Parameters.AddWithValue("@FNAME", txtEditAddBen.Text)
                cmd.Parameters.AddWithValue("@REL", txtEditAddBenRel.Text)
                cmd.Parameters.AddWithValue("@AGE", txtEditAddBenAge.Text)
                cmd.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox("Operation field")
            Finally
                conn.Close()
            End Try
            MsgBox(currentBen)
        Else
            MsgBox("Limit reached.")
            txtEditAddBen.Enabled = False
            txtEditAddBenAge.Enabled = False
            txtEditAddBenRel.Enabled = False
            btnEditAddBen.Enabled = False
        End If
        beneficiariesRecord()
    End Sub
    Private Sub btnEditNext_Click(sender As Object, e As EventArgs) Handles btnEditNext.Click
        tabEditMember.SelectedTab = other
        other.Enabled = True
        beneficiariesRecord()
        countBen()
    End Sub

    Private Sub btnEditBack_Click(sender As Object, e As EventArgs) Handles btnEditBack.Click
        tabEdit.Hide()
        pnlEmployee.Show()
    End Sub

    Private Sub tabEmployee_Click(sender As Object, e As EventArgs) Handles tabEmployee.Click
        pnlEmployee.Visible = True
        tabEdit.Visible = False
    End Sub


    Private Sub pickOffice_SelectedIndexChanged(sender As Object, e As EventArgs) Handles pickOffice.SelectedIndexChanged
        ' viewMembers("select users.id, concat(first_name, ' ', middle_name, ' ', last_name) as full_name, office, position, employment_status, 
        '                                    email from users left join user_info on users.id = user_info.user_id where office=" & pickOffice.SelectedItem)
        If pickOffice.SelectedIndex = 0 Then
            viewMembers("select users.id, concat(first_name, ' ', middle_name, ' ', last_name) as full_name, office, position, employment_status, 
                                            email from users left join user_info on users.id = user_info.user_id")
        Else
            dgMembers.Rows.Clear()
            Try
                conn.Open()
                Dim cmd As New MySqlCommand("select users.id, concat(first_name, ' ', middle_name, ' ', last_name) as full_name, office, position, employment_status, 
                                            email from users left join user_info on users.id = user_info.user_id where office=@OFC", conn)
                cmd.Parameters.AddWithValue("@OFC", pickOffice.SelectedItem)
                rid = cmd.ExecuteReader
                While rid.Read
                    dgMembers.Rows.Add(rid.Item("id"), rid.Item("full_name"), rid.Item("office"), rid.Item("position"), rid.Item("employment_status"), rid.Item("email"))
                End While
            Catch ex As Exception
                MsgBox("doesn't work lmao")
            Finally
                conn.Close()
            End Try
        End If
    End Sub

    Private Sub dgMembers_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgMembers.CellClick
        selectedId = dgMembers.CurrentRow.Cells(0).Value.ToString()
        If e.ColumnIndex = 7 AndAlso e.RowIndex >= 0 Then '----------------FOR EDIT

            Dim selectedId As Integer = dgMembers.CurrentRow.Cells(0).Value.ToString()
            other.Enabled = False
            tabEdit.Show()
            pnlEmployee.Hide()
            Dim locateProject As String = My.Application.Info.DirectoryPath
            Dim indext As Integer = locateProject.IndexOf("bin\Debug\net6.0-windows")
            Dim location As String = locateProject.Substring(0, indext)
            Try
                conn.Open()
                Dim cmd As New MySqlCommand("select * from users left join user_info on users.id = user_info.user_id where users.id=@ID", conn)
                cmd.Parameters.AddWithValue("@ID", selectedId)
                rid = cmd.ExecuteReader
                While rid.Read
                    If File.Exists(location & "\Resources\user_profile\" & rid.GetString("image")) Then
                        pBoxEditProfile.BackgroundImage = Image.FromFile(location & "\Resources\user_profile\" & rid.GetString("image"))
                    Else
                        pBoxEditProfile.BackgroundImage = Nothing
                    End If
                    txtEditUsername.Text = rid.GetString("username")
                    txtEditPw.Text = rid.GetString("password")
                    txtEditFname.Text = rid.GetString("first_name")
                    txtEditMname.Text = rid.GetString("middle_name")
                    txtEditLname.Text = rid.GetString("last_name")
                    txtEditNumber.Text = rid.GetString("contact")
                    txtEditAddress.Text = rid.GetString("address")
                    txtEditEducation.Text = rid.GetString("educational")
                    txtEditEmail.Text = rid.GetString("email")
                    pickEditOffice.Text = rid.GetString("office")
                    pickEditStatus.Text = rid.GetString("employment_status")
                    pickEditPosition.Text = rid.GetString("position")
                    pickEditComm.Text = rid.GetString("committee")
                    If rid.GetString("is_admin") = 1 Then
                        pickEditUserStat.Text = "Administrator"
                    Else
                        pickEditUserStat.Text = "Default"
                    End If
                End While
            Catch ex As Exception
            Finally
                conn.Close()
            End Try
        ElseIf e.ColumnIndex = 8 AndAlso e.RowIndex >= 0 Then '-------------FOR DELETE
            Dim result As DialogResult = MessageBox.Show("Are you sure you want to delete" & dgMembers.CurrentRow.Cells(1).Value.ToString() & "?", "Confirmation", MessageBoxButtons.YesNo)

            If result = DialogResult.Yes Then
                Dim selectedId As Integer = dgMembers.CurrentRow.Cells(0).Value.ToString()
                Try
                    conn.Open()
                    Dim cmd As New MySqlCommand("delete from users where id=@ID;
                                                    delete from user_info where user_id=@ID;
                                                        delete from beneficiaries where user_id=@ID", conn)
                    cmd.Parameters.AddWithValue("@ID", selectedId)
                    cmd.ExecuteNonQuery()
                Catch ex As Exception
                Finally
                    conn.Close()
                End Try
                MessageBox.Show("Deleted Successfully!")
            End If

        End If
        viewMembers("select users.id, concat(first_name, ' ', middle_name, ' ', last_name) as full_name, office, position, employment_status, 
                                            email from users left join user_info on users.id = user_info.user_id")
    End Sub
    Private Sub txtSearch_TextChanged(sender As Object, e As EventArgs) Handles txtSearch.TextChanged
        viewMembers("select users.id, concat(first_name, ' ', middle_name, ' ', last_name) as full_name, office, position, employment_status, 
                                            email from users left join user_info on users.id = user_info.user_id where first_name like '%" & txtSearch.Text & "%' or
                                            middle_name like '%" & txtSearch.Text & "%' or last_name like '%" & txtSearch.Text & "%' or office like '%" & txtSearch.Text &
                                    "%' or position like '%" & txtSearch.Text & "%' or employment_status like '%" & txtSearch.Text & "%' or email like '%" &
                                     txtSearch.Text & "%'")
    End Sub

    Private Sub dgMembers_CellContentClick_1(sender As Object, e As DataGridViewCellEventArgs) Handles dgMembers.CellContentClick

    End Sub

    Private Sub Guna2Button6_Click(sender As Object, e As EventArgs) Handles Guna2Button6.Click
        Dim AnswerYes As String
        AnswerYes = MsgBox("Are you sure you want to Log out", vbQuestion + vbYesNo, "User Repsonse")

        If AnswerYes = vbYes Then
            Form1.Show()
            Me.Hide()
        End If
    End Sub

    Private Sub btnEditUpdate_Click(sender As Object, e As EventArgs) Handles btnEditUpdate.Click
        Dim adminValue As Integer
        If pickEditUserStat.SelectedIndex = 0 Then
            adminValue = 1
        Else
            adminValue = 0
        End If
        Try
            conn.Open()
            Dim cmd As New MySqlCommand("update users set username=@USER, password=@PW, first_name=@FNAME, middle_name=@MNAME, last_name=@LNAME, position=@POS,
                                            is_admin=@ADMIN, updated_at=NOW() where id=@ID", conn)
            cmd.Parameters.AddWithValue("@USER", txtEditUsername.Text)
            cmd.Parameters.AddWithValue("@PW", txtEditPw.Text)
            cmd.Parameters.AddWithValue("@FNAME", txtEditFname.Text)
            cmd.Parameters.AddWithValue("@MNAME", txtEditMname.Text)
            cmd.Parameters.AddWithValue("@LNAME", txtEditLname.Text)
            cmd.Parameters.AddWithValue("@POS", pickEditPosition.SelectedItem)
            cmd.Parameters.AddWithValue("@ADMIN", adminValue)
            cmd.Parameters.AddWithValue("@ID", selectedId)
            cmd.ExecuteNonQuery()
            MsgBox("successfully updated.")
            pickOffice.SelectedIndex = 0
            tabEditMember.Hide()
            pnlEmployee.Show()
        Catch ex As Exception
            MsgBox("doesnt work update")
        Finally
            conn.Close()
        End Try
        viewMembers("select users.id, concat(first_name, ' ', middle_name, ' ', last_name) as full_name, office, position, employment_status, 
                                            email from users left join user_info on users.id = user_info.user_id")
    End Sub
End Class
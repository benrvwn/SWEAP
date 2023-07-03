﻿Imports TheArtOfDevHtmlRenderer.Adapters.Entities
Imports MySql.Data.MySqlClient
Imports System.IO
Imports Guna.UI2.WinForms
Imports OfficeOpenXml
Imports OfficeOpenXml.Style
Imports System.Text.RegularExpressions
Imports Microsoft.Office.Interop.Excel
Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Diagnostics.Metrics

Public Class admindash
    Dim conn As New MySqlConnection("server=172.30.206.128;port=3306;username=dswdSweap;password=druguser;database=sweap")
    Dim rid As MySqlDataReader
    Dim selectedId As Integer = 0
    Dim selectedBenId As Integer
    Dim currentBen As Integer
    Dim unionDue As Integer
    Dim countMembers As Integer
    Private Sub admindash_Load(sender As Object, e As EventArgs) Handles MyBase.Load '---------------AUTOLOAD
        getname()
        LoadChart()
        viewMembers("select users.id, concat(first_name, ' ', middle_name, ' ', last_name) as full_name, office, position, employment_status, 
                                            email from users left join user_info on users.id = user_info.user_id")
        viewMembersFundTransfer("Select users.id, concat(first_name, ' ', middle_name, ' ', last_name) as full_name, office, position, committee, 
                                            balance from users left join user_info On users.id = user_info.user_id")
        countmember()
        liveTimer.Start()

        If Guna2TabControl1.SelectedTab Is tabEmployee Then
            pnlEmployee.Visible = True ' Show the employeepanel
        Else
            tabEdit.Visible = False ' Hide the employeepanel
        End If

        Try
            conn.Open()
            Dim counter As Integer
            Dim cmd As New MySqlCommand("select * from contri_types", conn)
            rid = cmd.ExecuteReader
            While rid.Read
                dgContributions.Columns(3 + counter).HeaderText = rid.Item("alias")
                counter = counter + 1
            End While
        Catch ex As Exception
            MsgBox("header doesnt work")
        Finally
            conn.Close()
        End Try



        lblContri1.Text = dgContributions.Columns(3).HeaderText
        lblContri2.Text = dgContributions.Columns(4).HeaderText
        lblContri3.Text = dgContributions.Columns(5).HeaderText
        lblContri4.Text = dgContributions.Columns(6).HeaderText
        lblContri5.Text = dgContributions.Columns(7).HeaderText

        Try
            conn.Open()
            Dim cmd As New MySqlCommand("select concat(first_name, ' ', middle_name, ' ', last_name) as fullname, position, contributions.* from users left join contributions on users.id = contributions.user_id", conn)
            rid = cmd.ExecuteReader
            While rid.Read
                dgContributions.Rows.Add(rid.Item("user_id"), rid.Item("fullname"), rid.Item("position"), rid.Item("contribution1"), rid.Item("contribution2"), rid.Item("contribution3"), rid.Item("contribution4"), rid.Item("contribution5"))
            End While
        Catch ex As Exception
            MsgBox("contribution data fetching doesn't work")
        Finally
            conn.Close()
        End Try

    End Sub

    Private Sub Guna2Tabcontrol1_Click(sender As Object, e As EventArgs) Handles Guna2TabControl1.Click
        If Guna2TabControl1.SelectedTab Is tabEmployee Then
            pnlEmployee.Visible = True ' Show the employeepanel
        Else
            tabEdit.Visible = False ' Hide the employeepanel
        End If
    End Sub

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

    Public Sub countmember()
        Try
            conn.Open()
            Dim cmd As New MySqlCommand("Select count(*) as counts from users", conn)
            rid = cmd.ExecuteReader
            While rid.Read
                lblCntMember.Text = rid.GetInt32("counts")
            End While

        Catch ex As Exception
            MsgBox(ex.Message)
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

    Public Sub viewMembersFundTransfer(query) '-----------------FOR EMPLOYEES TABLE in Fund Transfer
        dgMembersFT.Rows.Clear()
        Try
            conn.Open()
            Dim cmd As New MySqlCommand(query, conn)
            rid = cmd.ExecuteReader
            While rid.Read
                dgMembersFT.Rows.Add(rid.Item("id"), rid.Item("full_name"), rid.Item("office"), rid.Item("position"), rid.Item("committee"), rid.Item("balance"))
            End While
        Catch ex As Exception
        Finally
            conn.Close()
        End Try
    End Sub



    Private Sub btnEditAddBen_Click(sender As Object, e As EventArgs) Handles btnEditAddBen.Click
        countBen()
        If currentBen < 5 Then
            Try
                If txtEditAddBen.Text = "" Or txtEditAddBenRel.Text = "" Or txtEditAddBenAge.Text = "" Then
                    MessageBox.Show("Please fill up the required field.", "Required Field", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    txtEditAddBen.BorderColor = Color.FromArgb(255, 0, 0)
                    txtEditAddBenRel.BorderColor = Color.FromArgb(255, 0, 0)
                    txtEditAddBenAge.BorderColor = Color.FromArgb(255, 0, 0)



                Else
                    conn.Open()
                    Dim cmd As New MySqlCommand("insert into beneficiaries(user_id, full_name, relationship, age) values(@ID, @FNAME, @REL, @AGE)", conn)
                    cmd.Parameters.AddWithValue("@ID", selectedId)
                    cmd.Parameters.AddWithValue("@FNAME", txtEditAddBen.Text)
                    cmd.Parameters.AddWithValue("@REL", txtEditAddBenRel.Text)
                    cmd.Parameters.AddWithValue("@AGE", txtEditAddBenAge.Text)
                    cmd.ExecuteNonQuery()

                    txtEditAddBen.Clear()
                    txtEditAddBenRel.Clear()
                    txtEditAddBenAge.Clear()
                End If
            Catch ex As Exception
                MsgBox("Operation field")
            Finally
                conn.Close()
            End Try
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
                MsgBox("doesn't Work")
            Finally
                conn.Close()
            End Try
        End If
    End Sub

    Private Sub txtSearch_TextChanged(sender As Object, e As EventArgs) Handles txtSearch.TextChanged
        viewMembers("select users.id, concat(first_name, ' ', middle_name, ' ', last_name) as full_name, office, position, employment_status, 
                                            email from users left join user_info on users.id = user_info.user_id where first_name like '%" & txtSearch.Text & "%' or
                                            middle_name like '%" & txtSearch.Text & "%' or last_name like '%" & txtSearch.Text & "%' or office like '%" & txtSearch.Text &
                                    "%' or position like '%" & txtSearch.Text & "%' or employment_status like '%" & txtSearch.Text & "%' or email like '%" &
                                     txtSearch.Text & "%'")
    End Sub


    Private Sub txtSearchbx_TextChanged(sender As Object, e As EventArgs) Handles txtSearchbx.TextChanged '--------Search name in Fundtransfer
        viewMembersFundTransfer("Select users.id, concat(first_name, ' ', middle_name, ' ', last_name) as full_name, office, position, committee, 
                                            balance from users left join user_info on users.id = user_info.user_id where  first_name like '%" & txtSearchbx.Text & "%' or
                                            middle_name like '%" & txtSearchbx.Text & "%' or last_name like '%" & txtSearchbx.Text & "%' or office like '%" & txtSearchbx.Text & "%' 
                                            or position like '%" & txtSearchbx.Text & "%' or committee like '%" & txtSearchbx.Text & "%' or balance like'%" & txtSearchbx.Text & "%' ")
    End Sub







    Private Sub Guna2Button6_Click(sender As Object, e As EventArgs) Handles Guna2Button6.Click
        Dim AnswerYes As String
        AnswerYes = MessageBox.Show("Are you sure you want to Log out", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If AnswerYes = vbYes Then
            Guna2TabControl1.SelectedTab = TabPage6
            Form2.Show()
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

    Private Sub dgSchedule_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgMembers.CellClick
        selectedId = dgMembers.CurrentRow.Cells(0).Value.ToString()
        If e.ColumnIndex = 6 AndAlso e.RowIndex >= 0 Then '----------------FOR EDIT

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
        ElseIf e.ColumnIndex = 7 AndAlso e.RowIndex >= 0 Then '-------------FOR DELETE
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




    Public Sub getname()
        Try
            conn.Open()
            Dim cmd As New MySqlCommand("select * from users  where id=@ID", conn)
            cmd.Parameters.AddWithValue("@ID", Form2.log_id)
            rid = cmd.ExecuteReader
            While rid.Read
                lblAdminName.Text = rid.GetString("first_name") + "!"
            End While
        Catch ex As Exception
        Finally
            conn.Close()
        End Try
    End Sub





    Private Sub dgBeneficiaries_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgBeneficiaries.CellClick
        If e.ColumnIndex = 4 AndAlso e.RowIndex >= 0 Then '-------------FOR DELETE
            Dim result As DialogResult = MessageBox.Show("are you sure you want to remove " & dgBeneficiaries.CurrentRow.Cells(1).Value.ToString() &
                                              " from beneficiaries?", "confirmation", MessageBoxButtons.YesNo)
            If result = DialogResult.Yes Then
                Try
                    conn.Open()
                    Dim cmd As New MySqlCommand("delete from beneficiaries where id=@id", conn)
                    cmd.Parameters.AddWithValue("@id", dgBeneficiaries.CurrentRow.Cells(0).Value.ToString())
                    cmd.ExecuteNonQuery()
                    MsgBox("record has been removed.")
                    beneficiariesRecord()
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
        End If
        beneficiariesRecord()
    End Sub

    'Key Press lang to pare!
    Private Shared Sub txtEditFname_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtEditFname.KeyPress
        If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsLetter(e.KeyChar) AndAlso Not Char.IsWhiteSpace(e.KeyChar) AndAlso Not Char.IsPunctuation(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub btnEmBack2_Click(sender As Object, e As EventArgs) Handles btnEmBack2.Click
        tabEditMember.SelectedTab = personal
    End Sub

    Private Sub txtEditMname_KeyPress(sender As Object, e As EventArgs)
        txtEditFname_KeyPress(sender, e)
    End Sub

    Private Sub txtEditLname_KeyPress(sender As Object, e As EventArgs)
        txtEditFname_KeyPress(sender, e)
    End Sub

    ''Email validation
    Private Function IsValidEmail(email As String) As Boolean
        Dim emailRegex As New Regex("^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$")
        Return emailRegex.IsMatch(email)
    End Function

    Private Sub txtEmail_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtEditEmail.Validating
        Dim inputEmail As String = txtEditEmail.Text.Trim()

        If txtEditEmail.Text = "" Then
            txtEditEmail.Text = txtEditEmail.Text
        ElseIf Not IsValidEmail(inputEmail) Then
            MessageBox.Show("Invalid email address." & vbCrLf & "Please enter a valid email address.", "INFORMATION", MessageBoxButtons.OK, MessageBoxIcon.Hand)
            e.Cancel = True

        End If
    End Sub

    '' txt type number only
    Private Sub txtEditNumber_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtEditNumber.KeyPress
        If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    '------------------------------------------------------CONTRIBUTIONS TAB-----------------------------------------------------------
    Private Sub btnEditContri_Click(sender As Object, e As EventArgs) Handles btnEditContri.Click
        Loan.tabconLoan.SelectedTab = Loan.contribution
        Loan.Show()
        Me.Close()
    End Sub

    Private Sub btnLoan_Click(sender As Object, e As EventArgs) Handles btnLoan.Click
        Loan.tabconLoan.SelectedTab = Loan.viewLoan
        Loan.Show()
        Me.Close()
    End Sub

    Private Sub liveTimer_Tick(sender As Object, e As EventArgs) Handles liveTimer.Tick '--------------------TIMER

        Dim timezone As TimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("singapore standard time")
        Dim currenttime As DateTime = TimeZoneInfo.ConvertTime(DateTime.Now, timezone)
        Dim currentdate As DateTime = currenttime
        Dim remainer As Integer
        lblDateTime.Text = currentdate.Hour & " : " & currentdate.Minute & " : " & currentdate.Second
        lblTime.Text = "AS OF " & currentdate.Date

        Try
            conn.Open()
            Dim cmd As New MySqlCommand("select sum(contribution1) as contri1, sum(contribution2) as contri2, sum(contribution3) as contri3,
                                        sum(contribution4) as contri4, sum(contribution5) as contri5 from contributions", conn)
            rid = cmd.ExecuteReader
            While rid.Read
                lblContri1Total.Text = rid.Item("contri1")
                lblContri2Total.Text = rid.Item("contri2")
                lblContri3Total.Text = rid.Item("contri3")
                lblContri4Total.Text = rid.Item("contri4")
                lblContri5Total.Text = rid.Item("contri5")
            End While
        Catch ex As Exception
            MsgBox("contribution dashboard doesn't work.")
        Finally
            conn.Close()
        End Try
    End Sub

    Private Sub Guna2Button4_Click(sender As Object, e As EventArgs) Handles Guna2Button4.Click
        Guna2TabControl1.SelectedTab = tabEmployee
    End Sub


    Sub LoadChart()
        With ChartMemberCount
            .Series.Clear()
            .Series.Add("Series1")
        End With

        Dim da As New MySqlDataAdapter("select office, sum(contribution1) + sum(contribution2) + sum(contribution3) + sum(contribution4) + sum(contribution5) as contri from contributions left join user_info on
user_info.user_id = contributions.user_id group by office", conn)
        Dim ds As New DataSet

        da.Fill(ds, "Members")
        ChartMemberCount.DataSource = ds.Tables("Members")
        Dim series1 As Series = ChartMemberCount.Series("Series1")
        series1.ChartType = SeriesChartType.Pie

        series1.Name = "MEMBERS"

        With ChartMemberCount
            .Series(0)("PieLabelStyle") = "Outside"
            .Series(0).BorderWidth = 1
            .Series(0).BorderColor = System.Drawing.Color.Black

            .Series(series1.Name).XValueMember = "office"
            .Series(series1.Name).YValueMembers = "contri"
            .Series(0).LabelFormat = "{#,##0}"
            .ChartAreas(0).Area3DStyle.Enable3D = True
            .Series(0).IsValueShownAsLabel = True

        End With
    End Sub

    Private Sub pickSex_SelectedIndexChanged(sender As Object, e As EventArgs) Handles pickSex.SelectedIndexChanged
        If pickSex.SelectedIndex = 0 Then
            viewMembers("select users.id, concat(first_name, ' ', middle_name, ' ', last_name) as full_name, office, position, employment_status, 
                                            email from users left join user_info on users.id = user_info.user_id")
        Else
            dgMembers.Rows.Clear()
            Try
                conn.Open()
                Dim cmd As New MySqlCommand("select users.id, concat(first_name, ' ', middle_name, ' ', last_name) as full_name, office, position, employment_status, 
                                            email from users left join user_info on users.id = user_info.user_id where sex=@sex", conn)
                cmd.Parameters.AddWithValue("@sex", pickSex.SelectedItem)
                rid = cmd.ExecuteReader
                While rid.Read
                    dgMembers.Rows.Add(rid.Item("id"), rid.Item("full_name"), rid.Item("office"), rid.Item("position"), rid.Item("employment_status"), rid.Item("email"))
                End While
            Catch ex As Exception
                MsgBox("doesn't Work")
            Finally
                conn.Close()
            End Try
        End If
    End Sub

    Private Sub dgMembersFT_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgMembersFT.CellClick '---------- To display data in transfer fund

        If e.ColumnIndex = 6 AndAlso e.RowIndex >= 0 AndAlso TypeOf dgMembersFT.Rows(e.RowIndex).Cells(e.ColumnIndex) Is DataGridViewImageCell Then

            Dim rowData As DataGridViewRow = dgMembersFT.Rows(e.RowIndex)

            Dim Value1 As String = rowData.Cells(0).Value.ToString()
            Dim Value2 As String = rowData.Cells(1).Value.ToString()
            Dim Value3 As String = rowData.Cells(5).Value.ToString()


            lblUserID.Text = Value1
            txtName.Text = Value2
            lblBalance.Text = Value3

            If Value3 = "" Then
                lblBalance.Text = "0"
            End If
        End If


    End Sub

    Private Sub bttnTransferFund_Click(sender As Object, e As EventArgs) Handles bttnTransferFund.Click
        Dim labelData As String = txtAmount.Text
        Dim labelname As String = txtName.Text
        Dim message As String = "Are you sure you want to add fund amounting " & labelData & " to account name: " & labelname & "?"


        Dim result As DialogResult = MessageBox.Show(message, "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If result = DialogResult.Yes Then

            Dim balance As Integer = lblBalance.Text
            Dim addFund As Integer = txtAmount.Text
            Dim sum As Integer = balance + addFund
            Try
                conn.Open()
                Dim cmd As New MySqlCommand("update users set balance = @balance where id=@id ", conn)

                cmd.Parameters.AddWithValue("@balance", sum)
                cmd.Parameters.AddWithValue("@id", lblUserID.Text)


                cmd.ExecuteNonQuery()
                MessageBox.Show("Fund transferred successfully!", "SUCCESSFULL", MessageBoxButtons.OK, MessageBoxIcon.Information)
                txtName.Clear()
                txtAmount.Clear()
                lblBalance.Text = "__"
                lblUserID.Text = "__"
            Catch ex As Exception
                MessageBox.Show("Fund transfer failed!", "FAILED", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Finally
                conn.Close()
            End Try
            viewMembersFundTransfer("Select users.id, concat(first_name, ' ', middle_name, ' ', last_name) as full_name, office, position, committee, 
                                            balance from users left join user_info On users.id = user_info.user_id")
        End If
    End Sub


End Class
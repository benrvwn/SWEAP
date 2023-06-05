﻿Imports TheArtOfDevHtmlRenderer.Adapters.Entities
Imports MySql.Data.MySqlClient
Public Class AdminDashboard
    Dim conn As New MySqlConnection("server=172.30.192.29;port=3306;username=sweapp;password=druguser;database=sweap")
    Dim rid As MySqlDataReader
    '-------------------------------FUNCTIONSS--------------------------------------------------------------------------------------
    Public Sub viewMembers(query)
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

    '-------------------------------------------------------------------------------------------------------------------------------
    Private Sub bttnDash_Click(sender As Object, e As EventArgs) Handles bttnDash.Click
        pnlDas.BackColor = Color.DarkRed
        pnlEm.BackColor = Color.Transparent
        pnlFundT.BackColor = Color.Transparent

        pnlDashboard.Visible = True
        pnlEmployees.Visible = False
        pnlFundTransfer.Visible = False

        'For button Dashboard design
        bttnDash.FillColor = Color.FromArgb(251, 197, 49)
        bttnDash.ForeColor = Color.Black
        bttnDash.Image = My.Resources.dashboard__1_
        bttnDash.BorderColor = Color.Black


        'For button Employee design
        bttnEmpl.FillColor = Color.DarkRed
        bttnEmpl.ForeColor = Color.White
        bttnEmpl.Image = My.Resources.employees
        bttnEmpl.BorderColor = Color.White


        'For button Fund Transfer design
        bttnFund.FillColor = Color.DarkRed
        bttnFund.ForeColor = Color.Wheat
        bttnFund.Image = My.Resources.transfer
        bttnFund.BorderColor = Color.White
    End Sub

    Private Sub bttnEmpl_Click(sender As Object, e As EventArgs) Handles bttnEmpl.Click
        pnlDas.BackColor = Color.Transparent
        pnlEm.BackColor = Color.DarkRed
        pnlFundT.BackColor = Color.Transparent

        pnlDashboard.Visible = False
        pnlEmployees.Visible = True
        pnlFundTransfer.Visible = False

        'For button Dashboard design
        bttnDash.FillColor = Color.Transparent
        bttnDash.ForeColor = Color.White
        bttnDash.Image = My.Resources.dash
        bttnDash.BorderColor = Color.White


        'For button Employee design
        bttnEmpl.FillColor = Color.FromArgb(251, 197, 49)
        bttnEmpl.ForeColor = Color.Black
        bttnEmpl.Image = My.Resources.company
        bttnEmpl.BorderColor = Color.Black


        'For button Fund Transfer design
        bttnFund.FillColor = Color.DarkRed
        bttnFund.ForeColor = Color.White
        bttnFund.Image = My.Resources.transfer
        bttnFund.BorderColor = Color.White

    End Sub

    Private Sub bttnFund_Click(sender As Object, e As EventArgs) Handles bttnFund.Click
        pnlDas.BackColor = Color.Transparent
        pnlEm.BackColor = Color.Transparent
        pnlFundT.BackColor = Color.DarkRed

        pnlDashboard.Visible = False
        pnlEmployees.Visible = False
        pnlFundTransfer.Visible = True

        'For button Dashboard design
        bttnDash.FillColor = Color.Transparent
        bttnDash.ForeColor = Color.Black
        bttnDash.Image = My.Resources.dashboard__1_
        bttnDash.BorderColor = Color.Black


        'For button Employee design
        bttnEmpl.FillColor = Color.Transparent
        bttnEmpl.ForeColor = Color.Black
        bttnEmpl.Image = My.Resources.company
        bttnEmpl.BorderColor = Color.Black


        'For button Fund Transfer design
        bttnFund.FillColor = Color.FromArgb(251, 197, 49)
        bttnFund.ForeColor = Color.Black
        bttnFund.Image = My.Resources.fund__1_
        bttnFund.BorderColor = Color.Black
    End Sub

    Private Sub pnlFundTransfer_Paint(sender As Object, e As PaintEventArgs) Handles pnlFundTransfer.Paint

    End Sub

    Private Sub Guna2Button3_Click(sender As Object, e As EventArgs) Handles Guna2Button3.Click

    End Sub

    Private Sub AdminDashboard_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        viewMembers("select users.id, concat(first_name, ' ', middle_name, ' ', last_name) as full_name, office, position, employment_status, 
                                            email from users left join user_info on users.id = user_info.user_id")
    End Sub

    Private Sub dgMembers_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgMembers.CellContentClick
        btnEditMember.Enabled = True
    End Sub

    Private Sub txtSearch_TextChanged(sender As Object, e As EventArgs) Handles txtSearch.TextChanged

        viewMembers("select users.id, concat(first_name, ' ', middle_name, ' ', last_name) as full_name, office, position, employment_status, 
                                            email from users left join user_info on users.id = user_info.user_id where first_name like '%" & txtSearch.Text & "%' or
                                            middle_name like '%" & txtSearch.Text & "%' or last_name like '%" & txtSearch.Text & "%' or office like '%" & txtSearch.Text &
                                            "%' or position like '%" & txtSearch.Text & "%' or employment_status like '%" & txtSearch.Text & "%' or email like '%" &
                                             txtSearch.Text & "%'")
    End Sub
End Class
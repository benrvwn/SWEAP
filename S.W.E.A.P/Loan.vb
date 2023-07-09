Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports DocumentFormat.OpenXml.Drawing.Charts
Imports DocumentFormat.OpenXml.Spreadsheet
Imports DocumentFormat.OpenXml.Wordprocessing
Imports OfficeOpenXml
Imports MySql.Data.MySqlClient
Imports DocumentFormat.OpenXml.Office.Word
Imports OfficeOpenXml.Style
Imports System.IO
Imports TheArtOfDevHtmlRenderer.Core.Utils
Imports System.Net.Sockets
Imports System.Text
Imports System.Net
Imports System.Globalization

'NAMING CONVENTION:
'btn = buttons
'pick = comboboxes
'txt = textboxes
'num = numberfields
'dg = datagrids
'date = datetimefields


Public Class Loan
    '-----------------------------------VARIABLE DECLARATION------------------------------------------
    Dim loanAmount As Decimal
    Dim annualInterestRate As Decimal
    Dim loanPeriodInYears As Integer
    Dim numberOfPaymentsPerYear As Integer
    Dim monthlyInterestRate As Decimal
    Dim totalNumberOfPayments As Integer
    Dim scheduledPayment As Decimal
    Dim extraPayment As Integer
    Dim payment As Integer
    Dim beginningBalance As Double
    Dim endBalance As Double
    Dim principal As Double
    Dim totalPayment As Double
    Dim interest As Double
    Dim CumuInterest As Double
    Dim selectedDate As DateTime
    Dim totalEarlyPayment As Integer
    Dim error_msg(0) As String '---used for validation, get's all the error messages if the validation function fails
    Dim random As Integer = 0 '---used to expand the error_msg array
    Dim errChecker As Integer = 0 '---used to check if the error_msg array is clear
    Dim message As String '---Validation messages, merging all the error messages from error_msg array to produce one message. Works along with errChecker variable
    Dim conn As New MySqlConnection(Form2.query) '---database connection
    Dim rid As MySqlDataReader '---database data reader
    Dim selectedId As Integer = 0
    Dim loanSchedId As Integer '---used in extract to excel loan to get some value to database
    '------------------------------------VARIABLE DECLARATION FOR CONTRIBUTIONS----------------------------------------------
    Dim updatedMonth As Integer '---get the updated month from database contributions
    Dim updatedYear As Integer '---Year
    Dim updatedWeek As Integer '---Week
    Dim updatedDay As Integer '---Day
    Dim query As String = "select user_id, concat(users.first_name, ' ', users.middle_name, ' ', users.last_name) as full_name, users.position, sum(contribution1) as contribution1,
                                        sum(contribution2) as contribution2, sum(contribution3) as contribution3, sum(contribution4) as contribution4, sum(contribution5) as contribution5, contributions.updated_at from contributions left join users
                                            on contributions.user_id = users.id group by contributions.user_id" '---used for contribution table
    '-----------------------------------------------END OF CONTRIBUTION'S VARIABLE-------------------------------------------

    '-----------------------------------END OF VARIABLE DECLARATION-------------------------------------------
    '--------------------------------------FUNCTIONS----------------------------------------------------------

    Public Function IsFileExists(filePath As String) As Boolean '---used to check if the file exist
        Return File.Exists(filePath)
    End Function

    Public Sub common_calculation() '---common calculation from add loan calculation
        totalPayment = scheduledPayment + extraPayment
        interest = beginningBalance * monthlyInterestRate
        principal = totalPayment - interest
        endBalance = beginningBalance - principal
        CumuInterest = CumuInterest + interest
    End Sub
    Public Sub common_process() '---common process of data passing from add loan calculation
        payment = 1
        selectedDate = dateStart.Value.Date
        extraPayment = numXtraP.Value
        loanAmount = numLamount.Value
        annualInterestRate = numAintRate.Value / 100
        loanPeriodInYears = numPyears.Value
        numberOfPaymentsPerYear = numPayYears.Value
        beginningBalance = loanAmount

        '--------------Convert the annual interest rate to monthly interest rate
        monthlyInterestRate = annualInterestRate / numberOfPaymentsPerYear

        '---------------Calculate the total number of payments
        totalNumberOfPayments = loanPeriodInYears * numberOfPaymentsPerYear

        '----------------Calculate the scheduled payment amount
        scheduledPayment = (monthlyInterestRate * loanAmount) / (1 - (1 + monthlyInterestRate) ^ (-totalNumberOfPayments))

        '------------------Round the scheduled payment amount to two decimal places
        scheduledPayment = Math.Round(scheduledPayment, 2)

        '--------------------------------------------OVERALL CALCULATION-----------------------------------------------------
    End Sub

    Public Sub validation(field, condition, msg)
        'USED FOR VALIDATIONS
        'parameters:
        'field = get the input field  and its value to validate / condition = value that validates the field / msg = message if the validation fails
        If field <= condition Then
            error_msg(random) = msg & vbNewLine
            random = random + 1
            ReDim Preserve error_msg(random) 'reseting arrays span every time another error message was added.
        End If
    End Sub
    Public Sub reset()
        'RESET ALL THE CURRENT VALUE OF COMPUTATIONS
        numXtraP.Value = 0
        numLamount.Value = 0
        numAintRate.Value = 0
        numPyears.Value = 0
        numPayYears.Value = 0
        loanAmount = 0
        btnSetSched.Enabled = True
        txtSnumberPayment.Text = ""
        txtSpayment.Text = ""
        txtActualNumPayment.Text = ""
        txtTotalEarlyPayment.Text = ""
        txtTotalInterest.Text = ""
        CumuInterest = 0
        dgSchedule.Rows.Clear()
        selectedId = 0
        txtLenderName.Text = ""
        txtName.Text = ""
    End Sub
    Public Sub mathing()
        'TO ROUNDOFF THE DECIMALS..
        beginningBalance = Math.Round(beginningBalance, 2)
        totalPayment = Math.Round(totalPayment, 2)
        interest = Math.Round(interest, 2)
        principal = Math.Round(principal, 2)
        endBalance = Math.Round(endBalance, 2)
        CumuInterest = Math.Round(CumuInterest, 2)
    End Sub

    Public Sub contriGrid(query) '--------------------FOR CONTRIBUTION TABLE
        dgContribution.Rows.Clear()
        Try
            conn.Open()
            Dim cmd As New MySqlCommand(query, conn)
            rid = cmd.ExecuteReader
            While rid.Read
                dgContribution.Rows.Add(rid.Item("user_id"), rid.Item("full_name"), rid.Item("position"), rid.Item("contribution1"), rid.Item("contribution2"), rid.Item("contribution3"), rid.Item("contribution4"), rid.Item("contribution5"), rid.Item("updated_at"))
            End While
        Catch ex As Exception
            MessageBox.Show("Fetching contribution table doesn't work. Function name contriGrid()", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            conn.Close()
        End Try
    End Sub

    Public Sub reset_contributions() '---------------------TO RESET OF CONTRIBUTIONS CLASS
        ' Clear the contributions array
        Array.Clear(contributions, 0, contributions.Length)

        Try
            conn.Open()
            ' Fetch contribution types from the database
            Dim cmd As New MySqlCommand("select * from contri_types", conn)
            Dim counter As Integer = 0
            rid = cmd.ExecuteReader
            While rid.Read
                ' Create a new instance of class_contribution and populate it with data from the database
                contributions(counter) = New class_contribution(rid.Item("contribution_name"), rid.Item("periodity"), rid.Item("amount"))

                ' Update the header text of the corresponding column in the DataGridView
                dgContribution.Columns(3 + counter).HeaderText = rid.Item("alias")

                ' Increment the counter for the next contribution type
                counter = counter + 1
            End While
        Catch ex As Exception
            MsgBox("Fetching of data failed from reset_contributions() function")
        Finally
            conn.Close()
        End Try

    End Sub

    Public Sub contriEditFields(status)
        'USED FOR ENABLING/DISABLING OF SOME FIELDS
        btnUpdateContriType.Enabled = status
        pickContriName.Enabled = status
        pickContriEditPeriod.Enabled = status
        txtNewContriName.Enabled = status
        numContriEditAmount.Enabled = status
    End Sub

    Public Sub forPickBox(input, query, selection)
        Try
            conn.Open() ' Open the database connection
            Dim cmd As New MySqlCommand(query, conn) ' Create a command with the provided query
            rid = cmd.ExecuteReader ' Execute the command and obtain a reader to retrieve the data
            While rid.Read ' Loop through each row of data
                input.Items.Add(rid.Item(selection)) ' Add the value of the selected column to the items of the input control
            End While
        Catch ex As Exception
            MessageBox.Show("Fetching data to combobox doesn't work", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            conn.Close() ' Close the database connection
        End Try
    End Sub

    Public Sub forHeader()

        Dim counter As Integer
        Try
            conn.Open()
            Dim cmd As New MySqlCommand("select alias from contri_types", conn)
            rid = cmd.ExecuteReader
            While rid.Read
                dgContriTotal.Columns(counter).HeaderText = rid.Item("alias")
                counter = counter + 1
            End While
        Catch ex As Exception
            MessageBox.Show("forHeader function doesn't work", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            conn.Close()
        End Try
    End Sub
    '------------------------------------END OF FUNCTIONS-----------------------------------------------------

    Private Sub btnSetSched_Click(sender As Object, e As EventArgs) Handles btnSetSched.Click '-------SET SCHEDULE BUTTON
        validation(numLamount.Value, 999, "* Loan can't be less than 1000.")
        validation(numAintRate.Value, 0, "* Anual Interest rate can't be less than or equal to 0.")
        validation(numPyears.Value, 0, "* Loan Period in years can't be less than or equal to 0.")
        validation(numPayYears.Value, 11, "* Number of payment per year can't be less than 12.")
        If selectedId = 0 Then
            ' Add an error message for blank lender name
            error_msg(random) = "* Lender name can't be blank." & vbNewLine
            random = random + 1
            ReDim Preserve error_msg(random)
        End If

        ' Concatenate all error messages into a single message
        While errChecker < error_msg.Length
            message = message & error_msg(errChecker)
            errChecker = errChecker + 1
        End While

        If message = "" Then
            ' Clear the rows in dgSchedule
            dgSchedule.Rows.Clear()

            ' Perform common process
            common_process()

            ' Perform common calculation
            common_calculation()

            While beginningBalance >= 0
                mathing()

                If beginningBalance < scheduledPayment Or beginningBalance < totalPayment Then
                    extraPayment = 0
                    totalPayment = beginningBalance
                    principal = totalPayment - interest
                    endBalance = 0
                End If

                ' Add a new row to dgSchedule with calculated values
                dgSchedule.Rows.Add(payment, selectedDate, "₱" & beginningBalance, "₱" & scheduledPayment, "₱" & extraPayment, "₱" & totalPayment, "₱" & principal, "₱" & interest, "₱" & endBalance, "₱" & CumuInterest)

                If endBalance = 0 Then
                    beginningBalance = 0
                    Exit While
                End If

                ' Update totalEarlyPayment and selectedDate for the next iteration
                totalEarlyPayment = totalEarlyPayment + extraPayment
                selectedDate = selectedDate.AddMonths(1)

                ' Update beginningBalance and perform common calculation
                beginningBalance = beginningBalance - principal
                common_calculation()

                payment = payment + 1
            End While

            '-------------------------------END OF CALCULATION-------------------------------

            '-------------------------------RESULT-------------------------------------------
            txtSnumberPayment.Text = numberOfPaymentsPerYear * loanPeriodInYears
            txtSpayment.Text = scheduledPayment
            txtActualNumPayment.Text = payment
            txtTotalEarlyPayment.Text = totalEarlyPayment
            txtTotalInterest.Text = CumuInterest
            txtName.Text = txtLenderName.Text
            '-------------------------------END OF RESULT-------------------------------------

            btnSetSched.Enabled = False
            btnApprove.Enabled = True
        Else
            ' Display error message in a MessageBox
            MessageBox.Show(message, "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Error)

            ' Reset variables and arrays
            errChecker = 0
            message = ""
            Array.Clear(error_msg, 0, error_msg.Length)
        End If

    End Sub

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        'RESET ALL THE FIELDS SO IS THE VALUE IN LOANS
        reset()
        btnApprove.Enabled = False '--disable the approve button everytime the reset clicked
    End Sub


    Private Sub btnSelectName_Click(sender As Object, e As EventArgs) Handles btnSelectName.Click
        'SHOW THE PANEL WHERE THERE'S A LIST OF USERS TO SELECT WHO'S GOING TO LOAN
        dgSelectEm.Rows.Clear()
        Try
            conn.Open()
            Dim cmd As New MySqlCommand("select concat(first_name, ' ', middle_name, ' ', last_name) as full_name, username, id from users", conn)
            rid = cmd.ExecuteReader
            While rid.Read
                dgSelectEm.Rows.Add(rid.Item("id"), rid.Item("full_name"), rid.Item("username"))
            End While
        Catch ex As Exception
            MessageBox.Show("Cannot connect to the database.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            conn.Close()
        End Try
        pnlSelectLender.Visible = True
    End Sub

    Private Sub btnBackPanel_Click(sender As Object, e As EventArgs) Handles btnBackPanel.Click
        'CLOSE THE PANEL OF USER'S LIST
        pnlSelectLender.Visible = False
    End Sub
    Private Sub btnApprove_Click(sender As Object, e As EventArgs) Handles btnApprove.Click
        'APPROVE THE CALCULATED SHCHEDULE OF LOAN AND STORE IT TO DATABASE
        'SAME PROCESS AS SET SCHEDULE BUTTON
        Dim result As DialogResult = MessageBox.Show("Do you want to proceed?", "Confirmation", MessageBoxButtons.YesNo)

        If result = DialogResult.Yes Then
            Dim random As New Random()
            Dim randomNum As Integer = random.Next(1, 801)
            CumuInterest = 0
            common_process()
            common_calculation()
            Try
                conn.Open()
                Dim cmd As New MySqlCommand("insert into loan_info(user_id, loans_id, loan_amount, anual_interest_rate, loan_period_years, 
                                            no_payments_per_year, start_date_of_loan, optional_xtra, status) values(@ID, @LID, @ALOAN, @ARATE, @LYEARS, @NOPYEAR, @SDATE, @XTRA, 'Ongoing');", conn)
                cmd.Parameters.AddWithValue("@ID", selectedId)
                cmd.Parameters.AddWithValue("@LID", randomNum)
                cmd.Parameters.AddWithValue("@ALOAN", numLamount.Value)
                cmd.Parameters.AddWithValue("@ARATE", numAintRate.Value)
                cmd.Parameters.AddWithValue("@LYEARS", numPyears.Value)
                cmd.Parameters.AddWithValue("@NOPYEAR", numPayYears.Value)
                cmd.Parameters.AddWithValue("@SDATE", dateStart.Value.Date)
                cmd.Parameters.AddWithValue("@XTRA", numXtraP.Value)
                cmd.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox("Loan_info doesn't work")
            Finally
                conn.Close()
            End Try
            While beginningBalance >= 0
                mathing()
                If beginningBalance < scheduledPayment Or beginningBalance < totalPayment Then
                    extraPayment = 0
                    totalPayment = beginningBalance
                    principal = totalPayment - interest
                    endBalance = 0
                End If

                Try
                    conn.Open()
                    Dim cmd As New MySqlCommand("INSERT INTO loans(user_id, loan_id, pmt_no, payment_date, beginning_balance, scheduled_payment, extra_payment,
                                                total_payment, principal, interest, ending_balance, cumulative_interest) VALUES(@ID, @LID,@PAYMENT,
                                                    @DATE, @BEGBAL, @SCHEDP, @XTRA, @TPAYMENT, @PRINCIPAL, @INTEREST, @ENDBAL, @CUMINTEREST)", conn)
                    cmd.Parameters.AddWithValue("@ALOAN", numLamount.Value)
                    cmd.Parameters.AddWithValue("@LID", randomNum)
                    cmd.Parameters.AddWithValue("@ID", selectedId)
                    cmd.Parameters.AddWithValue("@PAYMENT", payment)
                    cmd.Parameters.AddWithValue("@DATE", selectedDate)
                    cmd.Parameters.AddWithValue("@BEGBAL", beginningBalance)
                    cmd.Parameters.AddWithValue("@SCHEDP", scheduledPayment)
                    cmd.Parameters.AddWithValue("@XTRA", extraPayment)
                    cmd.Parameters.AddWithValue("@TPAYMENT", totalPayment)
                    cmd.Parameters.AddWithValue("@PRINCIPAL", principal)
                    cmd.Parameters.AddWithValue("@INTEREST", interest)
                    cmd.Parameters.AddWithValue("@ENDBAL", endBalance)
                    cmd.Parameters.AddWithValue("@CUMINTEREST", CumuInterest)
                    cmd.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox("Failed to insert data: " & ex.Message)
                Finally
                    conn.Close()
                End Try
                If endBalance = 0 Then
                    beginningBalance = 0
                    Exit While
                End If
                totalEarlyPayment = totalEarlyPayment + extraPayment
                selectedDate = selectedDate.AddMonths(1)
                beginningBalance = beginningBalance - principal
                common_calculation()
                payment = payment + 1

            End While

            MessageBox.Show("Loan added successfully!", "Response", MessageBoxButtons.OK, MessageBoxIcon.Information)
            reset()
        End If


    End Sub

    Private Sub Form2_Load_1(sender As Object, e As EventArgs) Handles MyBase.Load '------------AUTOLOAD
        'ALL THE INITIAL VALUES FUNCTIONS GOES HERE.
        pickContriOffice.SelectedIndex = 0 'set the selected office of contribution to all
        forHeader()
        btnLoanToExcel.Enabled = False
        contriEditFields(False)
        reset_contributions()
        contriGrid(query)
        forPickBox(pickContriOffice, "select office from user_info group by office", "office")
        'contriTimer.Start()
        Try '--fetch the latest date of contribution and store it to updated variables to be triggered to the contribute distribution function
            conn.Open()
            Dim cmd As New MySqlCommand("select month(updated_at) as month, year(updated_at) as year, dayofyear(updated_at) / 7 as week, day(updated_at) as day from contributions order by updated_at DESC limit 1", conn)
            rid = cmd.ExecuteReader
            While rid.Read
                updatedMonth = rid.GetInt32("month")
                updatedYear = rid.GetInt32("year")
                updatedWeek = rid.GetInt32("week")
                updatedDay = rid.GetInt32("day")
            End While
        Catch ex As Exception
            MessageBox.Show("Fetching updated date from the database doesn't work", "Response")
        Finally
            conn.Close()
        End Try
        dgSelectEm.Rows.Clear()

        'GET THE MEMBER'S DATA FROM THE DATABASE TO DATAGRID EMPLOYEE'S LIST ON LOAN PAGE
        Try
            conn.Open()
            Dim cmd As New MySqlCommand("select concat(first_name, ' ', middle_name, ' ', last_name) as full_name, username, id from users", conn)
            rid = cmd.ExecuteReader
            While rid.Read
                dgEmList.Rows.Add(rid.Item("id"), rid.Item("full_name"), rid.Item("username"))
            End While
        Catch ex As Exception
            MsgBox("doesn't work lmao")
        Finally
            conn.Close()
        End Try

        'GET THE CONTRIBUTIONS FROM THE DATABASE
        Try
            conn.Open()
            Dim cmd As New MySqlCommand("select sum(contribution1) as contri1, sum(contribution2) as contri2, sum(contribution3) as contri3, sum(contribution4) as contri4, sum(contribution5) as contri5 from contributions", conn)
            rid = cmd.ExecuteReader
            While rid.Read
                dgContriTotal.Rows.Add(rid.Item("contri1"), rid.Item("contri2"), rid.Item("contri3"), rid.Item("contri4"), rid.Item("contri5"))
            End While
        Catch ex As Exception
            MsgBox("contriTotal doesn't work")
        Finally
            conn.Close()
        End Try

        btnApprove.Enabled = False
        pnlSelectLender.Visible = False
        Guna2Button1.Enabled = False '-------------LOCK BUTTON
    End Sub

    Private Sub Guna2CircleButton1_Click(sender As Object, e As EventArgs) Handles Guna2CircleButton1.Click
        'CLOSE THE LOAN/CONTRIBUTIONS MODULE AND GO BACK TO ADMIN DASHBOARD
        admindash.Show()
        Me.Close()
    End Sub

    Private Sub dgSelectEm_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgSelectEm.CellClick
        'TRIGGERS IF THE 3RD INDEX IS CLICKED, USED TO SELECT LENDER
        If e.ColumnIndex = 3 AndAlso e.RowIndex >= 0 Then '----------------TO SELECT
            If String.IsNullOrEmpty(dgSelectEm.CurrentRow.Cells(0).Value.ToString()) = False Then
                selectedId = dgSelectEm.CurrentRow.Cells(0).Value.ToString()
                txtLenderName.Text = dgSelectEm.CurrentRow.Cells(1).Value.ToString()
                pnlSelectLender.Visible = False
            End If
        End If
    End Sub

    '-------------------------------------------------------VIEW LOANS---------------------------------------------------------------------------
    Private Sub dgEmList_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgEmList.CellClick
        Dim idSelect As Integer
        'TRIGGERS IF THE 3RD INDEX IS CLICKED, USED TO SELECT EMPLOYEES LOAN AND DISPLAYS TO ANOTHER DATAGRID BASED ON ITS ID.
        If e.ColumnIndex = 3 AndAlso e.RowIndex >= 0 Then '----------------TO SELECT
            btnLoanToExcel.Enabled = False
            dgLoans.Rows.Clear()
            dgLoanSchedule.Rows.Clear()
            idSelect = dgEmList.CurrentRow.Cells(0).Value.ToString()
            loanSchedId = idSelect
            Try
                conn.Open()
                Dim cmd As New MySqlCommand("select * from loan_info where user_id=@ID", conn)
                cmd.Parameters.AddWithValue("@ID", idSelect)
                rid = cmd.ExecuteReader
                While rid.Read
                    dgLoans.Rows.Add(rid.Item("loans_id"), rid.Item("loan_amount"), rid.Item("anual_interest_rate"), rid.Item("loan_period_years"), rid.Item("no_payments_per_year"), rid.Item("start_date_of_loan"), rid.Item("optional_xtra"), rid.Item("status"))
                End While
            Catch ex As Exception
                MsgBox("EW error")
            Finally
                conn.Close()
            End Try
        End If
    End Sub

    Private Sub dgLoans_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgLoans.CellClick
        Dim idSelect As Integer
        'TRIGGERS IF THE 9TH INDEX IS CLICKED, USED TO SELECT EMPLOYEES LOAN SCHEDULE AND DISPLAYS IT TO ANOTHERR DATAGRID BASED ON ITS ID.
        If e.ColumnIndex = 9 AndAlso e.RowIndex >= 0 Then '----------------TO SELECT
            idSelect = dgLoans.CurrentRow.Cells(0).Value.ToString()
            btnLoanToExcel.Enabled = True
            dgLoanSchedule.Rows.Clear()
            Try
                conn.Open()
                Dim cmd As New MySqlCommand("select * from loans where loan_id=@ID", conn)
                cmd.Parameters.AddWithValue("@ID", idSelect)
                rid = cmd.ExecuteReader
                While rid.Read
                    dgLoanSchedule.Rows.Add(rid.Item("pmt_no"), rid.Item("payment_date"), rid.Item("beginning_balance"), rid.Item("scheduled_payment"), rid.Item("extra_payment"), rid.Item("total_payment"), rid.Item("principal"), rid.Item("interest"), rid.Item("ending_balance"), rid.Item("cumulative_interest"))
                End While
            Catch ex As Exception
                MsgBox("EW error")
            Finally
                conn.Close()
            End Try
        End If
        'TRIGGERS IF THE 8TH INDEX IS CLICKED, USED TO MARK AS PAID TO THE LOAN
        If e.ColumnIndex = 8 AndAlso e.RowIndex >= 0 Then '----------------TO PAID

            If dgLoans.CurrentRow.Cells(7).Value.ToString() = "Ongoing" Then
                Dim result As DialogResult = MessageBox.Show("Is this loan paid already?" & vbNewLine & "Warning: You cannot change it back.", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                If result = DialogResult.Yes Then
                    idSelect = dgLoans.CurrentRow.Cells(0).Value.ToString()
                    Try
                        conn.Open()
                        Dim cmd As New MySqlCommand("update loan_info set status = 'Paid' where loans_id = @ID", conn)
                        cmd.Parameters.AddWithValue("@ID", idSelect)
                        cmd.ExecuteNonQuery()
                        MessageBox.Show("Record update succeeded!", "Response")
                    Catch ex As Exception
                        MessageBox.Show("Error updating status!", "Response")
                    Finally
                        conn.Close()
                    End Try
                    dgLoans.Rows.Clear()

                End If
            Else
                MessageBox.Show("Loan already marked as Paid.", "Response", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End If
    End Sub


    Private Sub btnLoanToExcel_Click(sender As Object, e As EventArgs) Handles btnLoanToExcel.Click '---------CREATE EXCEL FILE
        Dim filePath As String
        Dim fileOwner As String
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

        Using package As New ExcelPackage()
            Dim workbook As ExcelWorkbook = package.Workbook
            Dim worksheet As ExcelWorksheet = workbook.Worksheets.Add("Sheet1") ' Add a worksheet
            Dim counter As Integer = 13
            Dim totalXtraP As Integer

            worksheet.Cells("A2").Value = "ENTERED VALUES"
            worksheet.Cells("A3").Value = "Loan amount"
            worksheet.Cells("A4").Value = "Annual interest rate"
            worksheet.Cells("A5").Value = "Loan period in years"
            worksheet.Cells("A6").Value = "Loan amount"
            worksheet.Cells("A7").Value = "Annual interest rate"
            worksheet.Cells("A9").Value = "Loan period in years"


            For columnIndex As Integer = 1 To 8 ' Columns A to H
                worksheet.Column(columnIndex).Width = 50
            Next

            Dim columnRangeTops As ExcelRange = worksheet.Cells("A2:A9")
            columnRangeTops.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid
            columnRangeTops.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White)
            columnRangeTops.Style.Font.Color.SetColor(System.Drawing.Color.Black)
            columnRangeTops.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin
            columnRangeTops.Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black)
            columnRangeTops.Style.Font.Bold = True


            worksheet.Cells("G2").Value = "LOAN SUMMARY"
            worksheet.Cells("G3").Value = "Scheduled Payment"
            worksheet.Cells("G4").Value = "Scheduled number of payments"
            worksheet.Cells("G5").Value = "Actual number of payments"
            worksheet.Cells("G6").Value = "Total early payments"
            worksheet.Cells("G7").Value = "Total Interest"
            worksheet.Cells("G9").Value = "LENDER NAME"

            Dim columnRangeG As ExcelRange = worksheet.Cells("G2:G9")
            columnRangeG.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid
            columnRangeG.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White)
            columnRangeG.Style.Font.Color.SetColor(System.Drawing.Color.Black)
            columnRangeG.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin
            columnRangeG.Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black)
            columnRangeG.Style.Font.Bold = True

            Try
                conn.Open()
                Dim cmd As New MySqlCommand("select loan_info.*, concat(first_name, ' ', middle_name, ' ', last_name) as fullname, username from loan_info left join users on loan_info.user_id = users.id where user_id=@ID", conn)
                cmd.Parameters.AddWithValue("@ID", loanSchedId)
                rid = cmd.ExecuteReader
                While rid.Read
                    worksheet.Cells("B3").Value = rid.Item("loan_amount")
                    worksheet.Cells("B4").Value = rid.Item("anual_interest_rate")
                    worksheet.Cells("B5").Value = rid.Item("loan_period_years")
                    worksheet.Cells("B6").Value = rid.Item("no_payments_per_year")
                    worksheet.Cells("B8").Value = rid.Item("start_date_of_loan")
                    worksheet.Cells("B9").Value = rid.Item("optional_xtra")
                    worksheet.Cells("H4").Value = rid.Item("no_payments_per_year") * rid.Item("loan_period_years")
                    worksheet.Cells("H9").Value = rid.Item("fullname")
                    fileOwner = rid.Item("username")
                End While
            Catch ex As Exception
                MsgBox("doesntwork")
            Finally
                conn.Close()
            End Try

            Dim columnRangeB As ExcelRange = worksheet.Cells("B3:H9")
            columnRangeB.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid
            columnRangeB.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White)
            columnRangeB.Style.Font.Color.SetColor(System.Drawing.Color.Black)
            columnRangeB.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin
            columnRangeB.Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black)
            columnRangeB.Style.Font.Bold = True

            worksheet.Cells("H5").Value = dgLoanSchedule.RowCount
            worksheet.Cells("A11").Value = dgLoanSchedule.Columns(0).HeaderText
            worksheet.Cells("B11").Value = dgLoanSchedule.Columns(1).HeaderText
            worksheet.Cells("C11").Value = dgLoanSchedule.Columns(2).HeaderText
            worksheet.Cells("D11").Value = dgLoanSchedule.Columns(3).HeaderText
            worksheet.Cells("E11").Value = dgLoanSchedule.Columns(4).HeaderText
            worksheet.Cells("F11").Value = dgLoanSchedule.Columns(5).HeaderText
            worksheet.Cells("G11").Value = dgLoanSchedule.Columns(6).HeaderText
            worksheet.Cells("H11").Value = dgLoanSchedule.Columns(7).HeaderText
            worksheet.Cells("J11").Value = dgLoanSchedule.Columns(8).HeaderText
            worksheet.Cells("I11").Value = dgLoanSchedule.Columns(9).HeaderText

            Dim columnRange As ExcelRange = worksheet.Cells("A11:J11")
            columnRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid
            columnRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Green)
            columnRange.Style.Font.Color.SetColor(System.Drawing.Color.White)
            columnRange.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center
            columnRange.Style.Font.Bold = True


            ' Set column height
            worksheet.Row(11).Height = 20 ' Set the height of the header row

            For Each row As DataGridViewRow In dgLoanSchedule.Rows
                worksheet.Cells("A" & counter).Value = row.Cells(0).Value
                worksheet.Cells("B" & counter).Value = row.Cells(1).Value.ToString()
                worksheet.Cells("C" & counter).Value = row.Cells(2).Value
                worksheet.Cells("D" & counter).Value = row.Cells(3).Value
                worksheet.Cells("E" & counter).Value = row.Cells(4).Value
                worksheet.Cells("F" & counter).Value = row.Cells(5).Value
                worksheet.Cells("G" & counter).Value = row.Cells(6).Value
                worksheet.Cells("H" & counter).Value = row.Cells(7).Value
                worksheet.Cells("I" & counter).Value = row.Cells(8).Value
                worksheet.Cells("J" & counter).Value = row.Cells(9).Value
                totalXtraP = totalXtraP + row.Cells(4).Value
                counter = counter + 1
                worksheet.Cells("H7").Value = row.Cells(9).Value
                worksheet.Cells("H3").Value = row.Cells(3).Value
            Next
            worksheet.Cells("H6").Value = totalXtraP

            ' Set fill color and font color for all data rows
            Dim dataRange As ExcelRange = worksheet.Cells("A13:J" & (counter - 1))
            Dim fill As ExcelFill = dataRange.Style.Fill
            fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid
            fill.BackgroundColor.SetColor(System.Drawing.Color.White)
            dataRange.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin
            dataRange.Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black)
            Dim font As ExcelFont = dataRange.Style.Font
            font.Color.SetColor(System.Drawing.Color.Black)
            Dim location As String = My.Application.Info.DirectoryPath
            'Dim indext As Integer = locateProject.IndexOf("bin\Debug\net6.0-windows")
            'Dim location As String = locateProject.Substring(0, indext)

            For columnIndex As Integer = 1 To 2
                worksheet.Column(columnIndex).Width = 30
            Next
            For columnIndex As Integer = 3 To 6
                worksheet.Column(columnIndex).Width = 20
            Next
            For columnIndex As Integer = 7 To 8
                worksheet.Column(columnIndex).Width = 30
            Next
            For columnIndex As Integer = 9 To 10
                worksheet.Column(columnIndex).Width = 20
            Next



            filePath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & fileOwner & "_loan_Schedule.xlsx"

            If IsFileExists(filePath) Then
                Dim random As New Random()
                Dim randomNum As Integer = random.Next(1, 501)
                filePath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & fileOwner & "_loan_Schedule" & randomNum & ".xlsx"
                package.SaveAs(New System.IO.FileInfo(filePath))
                MsgBox("File saved to " & filePath)
            Else
                package.SaveAs(New System.IO.FileInfo(filePath))
                MsgBox("File saved to " & filePath)
            End If

        End Using
    End Sub

    '------------------------------------------------------------CONTRIBUTIONS--------------------------------------------------------------------

    Dim contriCounter As Integer = 0
    Public Shared contributions(4) As class_contribution

    Public Class class_contribution
        'This Class represents a contribution With its name, period, And amount. The constructor New initializes the instance variables
        'With the provided values.
        'The insertion method performs the database insertion Of the contribution. It opens the database connection, constructs the SQL query,
        'And uses parameters To prevent SQL injection. The query updates the contribution column With the specified amount multiplied by the
        'remainder value. Finally, the method executes the query And closes the database connection.
        Public contriName As String
        Public period As String
        Public amount As Integer

        Public Sub New(name As String, period As String, amount As Integer)
            ' Constructor for the class_contribution class
            ' Initializes the contribution name, period, and amount
            Me.contriName = name
            Me.period = period
            Me.amount = amount
        End Sub

        Public Sub insertion(remainder)
            ' Method to perform the insertion of contribution into the database
            Try
                Loan.conn.Open()
                Dim columnName As String = Me.contriName
                Dim query As String = "UPDATE contributions SET " & columnName & " = " & columnName & " + @AMOUNT, updated_at = now()"
                Dim cmd As New MySqlCommand(query, Loan.conn)
                cmd.Parameters.AddWithValue("@AMOUNT", Me.amount * remainder)
                cmd.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox("Insertion of contribution failed: " & ex.Message)
            Finally
                Loan.conn.Close()
            End Try
        End Sub
    End Class

    Private Sub btnUpdateContriType_Click(sender As Object, e As EventArgs) Handles btnUpdateContriType.Click
        'This code block performs an update operation on the contri_types table. It opens a database connection And
        'constructs an SQL query To update the Alias, amount, periodity, And updated_at columns. The @CN, @AMOUNT,
        '@PER, And @OCN parameters are used to provide values for the update operation.
        If pickContriName.Text = "" Then
            MessageBox.Show("No selected contribution name.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            Try
                conn.Open()
                Dim cmd As New MySqlCommand("update contri_types set alias = @CN, amount = @AMOUNT, periodity = @PER, updated_at = now() where alias=@OCN;", conn)
                cmd.Parameters.AddWithValue("@CN", txtNewContriName.Text)
                cmd.Parameters.AddWithValue("@AMOUNT", numContriEditAmount.Value)
                cmd.Parameters.AddWithValue("@PER", pickContriEditPeriod.SelectedItem)
                cmd.Parameters.AddWithValue("@OCN", pickContriName.SelectedItem)
                cmd.ExecuteNonQuery()
                MessageBox.Show("Update Succeeded!", "Response", MessageBoxButtons.OK, MessageBoxIcon.Information)
                dgContribution.Columns(pickContriName.SelectedIndex + 3).HeaderText = txtNewContriName.Text
            Catch ex As Exception
                MsgBox("Update doesn't work")
            Finally
                conn.Close()
            End Try
            contriEditFields(False)
            forHeader()
            btnOpenEdit.Enabled = True
            Guna2Button1.Enabled = False
        End If

    End Sub

    Private Sub btnOpenEdit_Click(sender As Object, e As EventArgs) Handles btnOpenEdit.Click
        pickContriName.Items.Clear()
        forPickBox(pickContriName, "select alias from contri_types", "alias")
        contriEditFields(True)
        btnOpenEdit.Enabled = False
        Guna2Button1.Enabled = True
    End Sub

    Private Sub pickContriName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles pickContriName.SelectedIndexChanged
        Try
            conn.Open() ' Open the database connection
            Dim cmd As New MySqlCommand("select * from contri_types where alias=@NAME", conn) ' Create a command to select data from the contri_types table
            cmd.Parameters.AddWithValue("@NAME", pickContriName.SelectedItem) ' Set the parameter value with the selected item
            rid = cmd.ExecuteReader ' Execute the command and obtain a reader to retrieve the data
            While rid.Read ' Loop through each row of data
                pickContriEditPeriod.SelectedItem = rid.Item("periodity") ' Set the selected item of pickContriEditPeriod to the value of the "periodity" column
                numContriEditAmount.Value = rid.Item("amount") ' Set the value of numContriEditAmount to the value of the "amount" column
            End While
        Catch ex As Exception
            MsgBox("Fetching data failed at pickContriName: " & ex.Message)
        Finally
            conn.Close() ' Close the database connection
        End Try

    End Sub

    Private Sub Guna2Button1_Click(sender As Object, e As EventArgs) Handles Guna2Button1.Click '-----------LOCK BUTTON FROM CONTRIBUTION
        contriEditFields(False)
        Guna2Button1.Enabled = False
        btnOpenEdit.Enabled = True
    End Sub

    Private Sub pickContriOffice_SelectedIndexChanged(sender As Object, e As EventArgs) Handles pickContriOffice.SelectedIndexChanged

        If pickContriOffice.SelectedItem = "All Members" Then
            contriGrid(query)
        Else
            dgContribution.Rows.Clear()
            Try
                conn.Open()
                Dim cmd As New MySqlCommand("select contributions.user_id, office, concat(users.first_name, ' ', users.middle_name, ' ', users.last_name) as full_name, users.position, sum(contribution1) as membership,
                    sum(contribution2) as union_due, sum(contribution3) as bereavement, sum(contribution4) as con4, sum(contribution5) as con5, contributions.updated_at from contributions left join users
                    on contributions.user_id = users.id left join user_info on contributions.user_id = user_info.user_id  where office = @OFFICE
                    group by contributions.user_id", conn)
                cmd.Parameters.AddWithValue("@OFFICE", pickContriOffice.SelectedItem)
                rid = cmd.ExecuteReader
                While rid.Read
                    dgContribution.Rows.Add(rid.Item("user_id"), rid.Item("full_name"), rid.Item("position"), rid.Item("membership"), rid.Item("union_due"), rid.Item("bereavement"), rid.Item("con4"), rid.Item("con5"), rid.Item("updated_at"))
                End While
            Catch ex As Exception
                MessageBox.Show("Connection failed.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                conn.Close()
            End Try
            dgContriTotal.Rows.Clear()
            Try
                conn.Open()
                Dim cmd As New MySqlCommand("select office, sum(contribution1) as contri1, sum(contribution2) as contri2, sum(contribution3) as contri3, sum(contribution4) as contri4, sum(contribution5) as contri5 from contributions left join user_info on contributions.user_id = user_info.user_id
                                            where office = @OFFICE", conn)
                cmd.Parameters.AddWithValue("@OFFICE", pickContriOffice.SelectedItem)
                rid = cmd.ExecuteReader
                While rid.Read
                    dgContriTotal.Rows.Add(rid.Item("contri1"), rid.Item("contri2"), rid.Item("contri3"), rid.Item("contri4"), rid.Item("contri5"))
                End While
            Catch ex As Exception
            Finally
                conn.Close()
            End Try
        End If



    End Sub

    Private Sub btnExtractContri_Click(sender As Object, e As EventArgs) Handles btnExtractContri.Click
        Dim filePath As String
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

        ' Create a new Excel package
        Using package As New ExcelPackage()
            Dim workbook As ExcelWorkbook = package.Workbook
            Dim worksheet As ExcelWorksheet = workbook.Worksheets.Add("Sheet1")
            Dim counter As Integer = 13

            ' Set cell values for headers and data in the worksheet
            worksheet.Cells("A2").Value = "OFFICE"
            worksheet.Cells("A3").Value = "TOTAL CONTRIBUTIONS"
            worksheet.Cells("A5").Value = dgContriTotal.Columns(0).HeaderText
            worksheet.Cells("A6").Value = dgContriTotal.Columns(1).HeaderText
            worksheet.Cells("A7").Value = dgContriTotal.Columns(2).HeaderText
            worksheet.Cells("A8").Value = dgContriTotal.Columns(3).HeaderText
            worksheet.Cells("A9").Value = dgContriTotal.Columns(4).HeaderText

            worksheet.Cells("C2").Value = dgContriTotal.Rows(0).Cells(0).Value
            worksheet.Cells("C5").Value = dgContriTotal.Rows(0).Cells(0).Value
            worksheet.Cells("C6").Value = dgContriTotal.Rows(0).Cells(1).Value
            worksheet.Cells("C7").Value = dgContriTotal.Rows(0).Cells(2).Value
            worksheet.Cells("C8").Value = dgContriTotal.Rows(0).Cells(3).Value
            worksheet.Cells("C9").Value = dgContriTotal.Rows(0).Cells(4).Value

            worksheet.Cells("A12").Value = dgContribution.Columns(0).HeaderText
            worksheet.Cells("B12").Value = dgContribution.Columns(1).HeaderText
            worksheet.Cells("C12").Value = dgContribution.Columns(2).HeaderText
            worksheet.Cells("D12").Value = dgContribution.Columns(3).HeaderText
            worksheet.Cells("E12").Value = dgContribution.Columns(4).HeaderText
            worksheet.Cells("F12").Value = dgContribution.Columns(5).HeaderText
            worksheet.Cells("G12").Value = dgContribution.Columns(6).HeaderText

            ' Iterate over each row in the DataGridView and populate the worksheet cells
            For Each row As DataGridViewRow In dgContribution.Rows
                worksheet.Cells("A" & counter).Value = row.Cells(0).Value
                worksheet.Cells("B" & counter).Value = row.Cells(1).Value
                worksheet.Cells("C" & counter).Value = row.Cells(2).Value
                worksheet.Cells("D" & counter).Value = row.Cells(3).Value
                worksheet.Cells("E" & counter).Value = row.Cells(4).Value
                worksheet.Cells("F" & counter).Value = row.Cells(5).Value
                worksheet.Cells("G" & counter).Value = row.Cells(6).Value
                counter = counter + 1
            Next

            ' Specify the file path to save the Excel file
            Dim location As String = My.Application.Info.DirectoryPath
            'Dim indext As Integer = locateProject.IndexOf("bin\Debug\net6.0-windows")
            'Dim location As String = locateProject.Substring(0, indext)

            filePath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & pickContriOffice.SelectedItem & "_Contribution.xlsx"

            ' Check if the file already exists, if so, add a random number to the file name
            If IsFileExists(filePath) Then
                Dim random As New Random()
                Dim randomNum As Integer = random.Next(1, 501)
                filePath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & pickContriOffice.SelectedItem & "_Contribution" & randomNum & ".xlsx"
                package.SaveAs(New System.IO.FileInfo(filePath))
                MessageBox.Show("File saved to " & filePath, "Response")
            Else
                package.SaveAs(New System.IO.FileInfo(filePath))
                MessageBox.Show("File saved to " & filePath, "Response")
            End If

        End Using
    End Sub

    Public Sub contriTrigger(upDate, currentDate, period)
        ' This method triggers contribution updates if the given date has changed
        Dim remainder As Integer
        If upDate <> currentDate Then
            remainder = currentDate - upDate
            For Each contribution As class_contribution In contributions
                If contribution.period = period Then
                    contribution.insertion(remainder)
                End If
            Next
            contriGrid(query)
        End If
    End Sub

    Private Sub btnUpToDateContri_Click(sender As Object, e As EventArgs) Handles btnUpToDateContri.Click

        ' Retrieve contribution data and perform updates
        pickContriOffice.SelectedIndex = 0
        contriGrid(query)

        Dim currentdate As DateTime
        Dim currentweek As Integer

        Try
            ' Get the current internet time from an NTP server
            Dim ntpServer As String = "pool.ntp.org"
            Dim ntpPort As Integer = 123

            Dim clientSocket As New UdpClient()
            Dim endPoint As New IPEndPoint(Dns.GetHostAddresses(ntpServer)(0), ntpPort)
            clientSocket.Connect(endPoint)

            Dim ntpData As Byte() = New Byte(47) {}
            ntpData(0) = &H1B ' Set Mode to Client

            clientSocket.Send(ntpData, ntpData.Length)

            Dim responseData As Byte() = clientSocket.Receive(endPoint)
            clientSocket.Close()

            Array.Reverse(responseData, 40, 4) ' Reverse byte order for timestamp

            Dim intPart As UInteger = BitConverter.ToUInt32(responseData, 40)
            Dim fracPart As UInteger = BitConverter.ToUInt32(responseData, 44)

            Dim milliseconds = (intPart * 1000) + ((fracPart * 1000) / &H100000000UL)

            Dim baseDateTime As New DateTime(1900, 1, 1)
            Dim networkDateTime As DateTime = baseDateTime.AddMilliseconds(milliseconds)

            ' Convert to Singapore Standard Time
            Dim singaporeTimeZone As TimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("Asia/Singapore")
            Dim singaporeDateTime As DateTime = TimeZoneInfo.ConvertTimeFromUtc(networkDateTime, singaporeTimeZone)
            currentdate = singaporeDateTime
            currentweek = currentdate.DayOfYear / 7

        Catch ex As Exception
            MessageBox.Show("An error occurred while getting the internet time: " & ex.Message)
        End Try
        'MsgBox(currentdate.Month & " " & currentdate.Year & " " & currentweek & " " & currentdate.Day & vbNewLine & updatedMonth & " " & updatedYear & " " & updatedWeek & " " & updatedDay)

        contriTrigger(updatedMonth, currentdate.Month, "Monthly")
        contriTrigger(updatedYear, currentdate.Year, "Annually")
        contriTrigger(updatedWeek, currentweek, "Weekly")
        contriTrigger(updatedDay, currentdate.Day, "Daily")

        ' Update the updatedMonth, updatedYear, updatedWeek, and updatedDay variables
        updatedMonth = currentdate.Month
        updatedYear = currentdate.Year
        updatedWeek = currentweek
        updatedDay = currentdate.Day

        MessageBox.Show("RECORD IS UP TO DATE " & vbNewLine & vbNewLine & currentdate, "Response", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
    '----------------------------------------------END OF CONTRIBUTIONS-----------------------------------------------------------


End Class


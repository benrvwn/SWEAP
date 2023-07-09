Imports System.IO
Imports MySql.Data.MySqlClient
Imports OfficeOpenXml
Imports OfficeOpenXml.Style
Imports System.Text.RegularExpressions
Imports Microsoft.VisualBasic.FileIO

Public Class user_dashboard
    '
    Dim conn As New MySqlConnection(Form2.query)
    ' MySqlConnection is used to establish a connection to a MySQL database.

    Dim dr As MySqlDataReader
    ' MySqlDataReader is used to retrieve data from a MySQL database.

    Dim sourceFilePath As String
    ' This variable is used to store the path of a source file.

    Dim getExtension As String
    ' This variable is used to store the file extension of a file.

    Dim location As String = My.Application.Info.DirectoryPath
    ' My.Application.Info.DirectoryPath retrieves the directory path of the current application.

    'Dim indext As Integer = locateProject.IndexOf("bin\Debug\net6.0-windows")
    ' IndexOf method is used to find the position of a specific string within another string.

    'Dim location As String = locateProject.Substring(0, indext)
    ' Substring method is used to extract a portion of a string based on the specified start and end indexes.

    Dim destinationPath As String = location & "\Others\images"
    ' This variable is used to store the destination path for the user profile.

    Dim destinationIconPath As String = location & "\Others\"
    ' This variable is used to store the destination path for icons.

    Dim dashPath As String = "dashboard (3).png"
    ' This variable stores the file name of the dashboard image.

    Dim profPath As String = "man.png"
    ' This variable stores the file name of the default profile image.

    Dim benefPath As String = "beneficiary (2).png"
    ' This variable stores the file name of the beneficiary image.

    Dim settingPath As String = "settings.png"
    ' This variable stores the file name of the settings image.

    Dim Home As String = "house (1).png"
    ' This variable stores the file name of the home image.

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        ' Timer1_Tick is an event handler that executes when the Timer1's interval has elapsed.

        Dim timezone As TimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("singapore standard time")
        ' TimeZoneInfo.FindSystemTimeZoneById is used to retrieve the time zone information for the specified time zone identifier.

        Dim currenttime As DateTime = TimeZoneInfo.ConvertTime(DateTime.Now, timezone)
        ' TimeZoneInfo.ConvertTime is used to convert the current time to the specified time zone.

        Dim currentdate As DateTime = currenttime
        ' This variable stores the current date and time.

        Dim remainer As Integer
        ' This variable is not used in the provided code.

        lblTime.Text = currentdate.Hour & " : " & currentdate.Minute & " : " & currentdate.Second
        ' lblTime.Text is set to display the current time in the format: "hour : minute : second".
    End Sub

    Private Sub user_dashboard_load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' user_dashboard_load is an event handler that executes when the form is loaded.

        Timer1.Start()
        ' The Timer1 is started to initiate the Timer1_Tick event.

        Get_info()
        ' A method or function named Get_info is called.

        DG_Load()
        ' A method or function named DG_Load is called.

        Dashboard()
        ' A method or function named Dashboard is called.

        conn.Open()
        Dim cmd As New MySqlCommand("select alias from contri_types", conn)
        dr = cmd.ExecuteReader()
        Dim rowIndex As Integer = 0
        While dr.Read() AndAlso rowIndex < 5
            If rowIndex = 0 Then
                lblContri1.Text = dr.GetString(0)
            ElseIf rowIndex = 1 Then
                lblContri2.Text = dr.GetString(0)
            ElseIf rowIndex = 2 Then
                lblContri3.Text = dr.GetString(0)
            ElseIf rowIndex = 3 Then
                lblContri4.Text = dr.GetString(0)
            ElseIf rowIndex = 4 Then
                lblContri5.Text = dr.GetString(0)
            End If

            rowIndex += 1
        End While
        conn.Close()

    End Sub


    'FETCHED Balance, Contribution1, and Contribution2 of the user.
    Public Sub Dashboard()
        Try
            conn.Open()  ' Opens a connection to the database.
            Dim cmd As New MySqlCommand("SELECT * FROM users 
                                LEFT JOIN contributions on users.id = contributions.user_id 
                                WHERE users.id = @id", conn)  ' Creates a new MySqlCommand object to retrieve data from the "users" table and calculate the sum of union dues and bereavement contributions using a LEFT JOIN with the "contributions" table. The query filters the results based on the user's id (Form2.log_id).
            cmd.Parameters.AddWithValue("@id", Form2.log_id) ' Passes the log_id from Form2 as a parameter to the query.
            dr = cmd.ExecuteReader ' Executes the query and populates the DataReader object "dr" with the results.
            While dr.Read ' While reading the results from the DataReader...
                If Not dr.IsDBNull(dr.GetOrdinal("balance")) Then ' Checks if the "balance" column is not null for the current row.
                    txtSaving.Text = dr.GetString("balance")  ' Sets the text of txtSaving TextBox to the value of the "balance" column.
                Else
                    txtSaving.Text = 0 ' If the "balance" column is null, sets the text of txtSaving TextBox to 0.
                End If

                If Not dr.IsDBNull(dr.GetOrdinal("contribution1")) Then ' checks if the "contribution1" column is not null for the current row.
                    txtContribution1.Text = dr.GetString("contribution1") ' sets the text of txtcontribution1 textbox to the value of the "contribution1" column.
                Else
                    txtContribution1.Text = 0 ' if the "contribution1" column is null, sets the text of txtcontribution1 textbox to 0.
                End If

                If Not dr.IsDBNull(dr.GetOrdinal("contribution2")) Then ' Checks if the "contribution2" column is not null for the current row.
                    txtContribution2.Text = dr.GetString("contribution2") ' Sets the text of txtContribution2 TextBox to the value of the "contribution2" column.
                Else
                    txtContribution2.Text = 0 ' If the "contribution3" column is null, sets the text of txtContribution2 TextBox to 0.
                End If

                If Not dr.IsDBNull(dr.GetOrdinal("contribution3")) Then ' Checks if the "contribution3" column is not null for the current row.
                    txtContribution3.Text = dr.GetString("contribution3") ' Sets the text of txtContribution3 TextBox to the value of the "contribution3" column.
                Else
                    txtContribution3.Text = 0 ' If the "contribution4" column is null, sets the text of txtContribution4 TextBox to 0.
                End If

                If Not dr.IsDBNull(dr.GetOrdinal("contribution4")) Then ' Checks if the "contribution4" column is not null for the current row.
                    txtContribution4.Text = dr.GetString("contribution4") ' Sets the text of txtContribution4 TextBox to the value of the "contribution4" column.
                Else
                    txtContribution4.Text = 0 ' If the "contribution5" column is null, sets the text of txtContribution5 TextBox to 0.
                End If

                If Not dr.IsDBNull(dr.GetOrdinal("contribution5")) Then ' Checks if the "contribution5" column is not null for the current row.
                    txtContribution5.Text = dr.GetString("contribution5") ' Sets the text of txtContribution5 TextBox to the value of the "contribution5" column.
                Else
                    txtContribution5.Text = 0 ' If the "contribution5" column is null, sets the text of txtContribution5 TextBox to 0.
                End If
            End While
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            conn.Close() ' Closes the database connection.
        End Try
    End Sub

    'LOG OUT-------------------------
    Private Sub btnLogOut_Click(sender As Object, e As EventArgs) Handles btnLogOut.Click
        Dim AnswerYes As String
        AnswerYes = MessageBox.Show("Are you sure you want to Log out?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If AnswerYes = vbYes Then
            Guna2TabControl1.SelectedTab = tabDashboard ' Selects the "tabDashboard" tab in the Guna2TabControl.
            Me.Hide() ' Hides the current form.
            Form2.Show() ' Shows the Form2.
            lblFromTitle.Text = "Home" ' Sets the text of lblFromTitle label to "Home".
            iconFromtitle.Image = Image.FromFile(destinationIconPath + Home) ' Sets the image of iconFromtitle PictureBox to the image located at the destinationIconPath concatenated with the "Home" file name.
        End If

    End Sub


    Private Sub Guna2TabControl1_Selected(sender As Object, e As TabControlEventArgs) Handles Guna2TabControl1.Selected
        If Guna2TabControl1.SelectedTab Is tabDashboard Then 'If the selected tab is "tabDashboard"...
            lblFromTitle.Text = "Dashboard" ' Sets the text of lblFromTitle label to "Dashboard"
            iconFromtitle.Image = Image.FromFile(destinationIconPath + dashPath) ' Sets the image of iconFromtitle PictureBox to the image located at the destinationIconPath concatenated with the dashPath file name.
            Dashboard() ' Calls the Dashboard() function to populate the dashboard information.
        ElseIf Guna2TabControl1.SelectedTab Is tabProfile Then ' If the selected tab is "tabProfile"...
            lblFromTitle.Text = "Profile" ' Sets the text of lblFromTitle label to "Profile".
            iconFromtitle.Image = Image.FromFile(destinationIconPath + profPath) ' Sets the image of iconFromtitle PictureBox to the image located at the destinationIconPath concatenated with the profPath file name.
            Get_info() ' Calls the Get_info() function to retrieve and display the user's profile information.
        ElseIf Guna2TabControl1.SelectedTab Is tabBeneficiary Then ' If the selected tab is "tabBeneficiary"...
            lblFromTitle.Text = "Beneficiary" ' Sets the text of lblFromTitle label to "Beneficiary".
            iconFromtitle.Image = Image.FromFile(destinationIconPath + benefPath) ' Sets the image of iconFromtitle PictureBox to the image located at the destinationIconPath concatenated with the benefPath file name.
            DG_Load() ' Calls the DG_Load() function to load and display the beneficiary data in the DataGridView.
        ElseIf Guna2TabControl1.SelectedTab Is tabSetting Then ' If the selected tab is "tabSetting"...
            lblFromTitle.Text = "Account Setting" ' Sets the text of lblFromTitle label to "Account Setting".
            iconFromtitle.Image = Image.FromFile(destinationIconPath + settingPath) ' Sets the image of iconFromtitle PictureBox to the image located at the destinationIconPath concatenated with the settingPath file name.
        Else ' For any other tab...
            lblFromTitle.Text = "Home" ' Sets the text of lblFromTitle label to "Home".
            iconFromtitle.Image = Image.FromFile(destinationIconPath + Home) ' Sets the image of iconFromtitle PictureBox to the image located at the destinationIconPath concatenated with the Home file name.
        End If
    End Sub


    'FETCHED ALL BENEFICIARY OF USER-------------------------------
    Public Sub DG_Load()
        BeneficiariesDGV.Rows.Clear()
        Try
            conn.Open() ' Opens a connection to the database.
            Dim cmd As New MySqlCommand("SELECT beneficiaries.* FROM users 
                                INNER JOIN beneficiaries ON users.id = beneficiaries.user_id
                                WHERE users.id = @ID", conn)  ' Creates a new MySqlCommand object to retrieve data from the "beneficiaries"
            ' table by joining it with the "users" table based on the user_id. The query filters the results based on the user's id (Form2.log_id).
            cmd.Parameters.AddWithValue("@ID", Form2.log_id)  'Passes the log_id from Form2 as a parameter to the query.
            dr = cmd.ExecuteReader() ' Executes the query and populates the DataReader object "dr" with the results.
            While dr.Read() ' While reading the results from the DataReader...
                BeneficiariesDGV.Rows.Add(dr.Item("user_id"), dr.Item("full_name"), dr.Item("age"), dr.Item("relationship"))
                ' Adds a new row to the BeneficiariesDGV DataGridView and populates it with the values from the current row of the DataReader.
            End While
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            conn.Close()  'Closes the database connection.
        End Try
    End Sub


    'FETCH ALL INFO OF THE USER------------------------
    Public Sub Get_info()
        Try
            conn.Open() ' Opens a connection to the database.
            Dim cmd As New MySqlCommand("SELECT *, CONCAT(first_name, ' ', middle_name, ' ', last_name) AS fullName FROM users 
                            INNER JOIN user_info ON users.id = user_info.user_id  
                            WHERE users.id = @ID", conn)  ' Creates a new MySqlCommand object to retrieve all information of the user by joining the "users" table with the "user_info" table
            ' based on the user_id. The query filters the results based on the user's id (Form2.log_id).
            cmd.Parameters.AddWithValue("@ID", Form2.log_id) ' Passes the log_id from Form2 as a parameter to the query.
            dr = cmd.ExecuteReader ' Executes the query and populates the DataReader object "dr" with the results.
            While dr.Read ' While reading the results from the DataReader...
                Dim imagepathResources As String = destinationPath + dr.GetString("image") ' Retrieves the image path from the result and concatenates it with the destinationPath to get the complete image path.

                If File.Exists(imagepathResources) Then ' If the image file exists at the specified path...
                    userProfile.Image = Image.FromFile(imagepathResources) ' Sets the image of the userProfile PictureBox to the image located at the imagepathResources.
                    user_Profile.BackgroundImage = Image.FromFile(imagepathResources) ' Sets the image of the user_Profile PictureBox to the image located at the imagepathResources.
                Else ' If the image file does not exist...
                    userProfile.Image = Nothing ' Clears the image of the userProfile PictureBox.
                    user_Profile.BackgroundImage = Nothing ' Clears the image of the user_Profile PictureBox.
                End If

                ' Determines the appropriate greeting based on the selected value of cmboSex ComboBox.
                If cmboSex.SelectedIndex = 0 Then
                    lblDateTime.Text = "Mr. " + dr.GetString("first_name") + " Your last log in was " + dr.GetString("last_logout")
                ElseIf cmboSex.SelectedIndex = 1 Then
                    lblDateTime.Text = "Ms. " + dr.GetString("first_name") + " Your last log in was " + dr.GetString("last_logout")
                ElseIf cmboSex.SelectedIndex = 2 Then
                    lblDateTime.Text = "Hi " + dr.GetString("first_name") + " Your last log in was " + dr.GetString("last_logout")
                End If

                ' Populates the various labels and text boxes with the retrieved information.
                Pfname.Text = dr.GetString("fullName")
                Padd.Text = dr.GetString("address")
                Pcntact.Text = dr.GetString("contact")
                Pemail.Text = dr.GetString("email")
                Pbdate.Text = dr.GetDateTime("birthdate").ToString()

                Peducational.Text = dr.GetString("educational")
                Pemployment.Text = dr.GetString("employment_status")
                Poffice.Text = dr.GetString("office")
                Pposition.Text = dr.GetString("position")
                Pcommittee.Text = dr.GetString("committee")
                PSex.Text = dr.GetString("sex")

                txtbxusername.Text = dr.GetString("username")
                txtbxpassword.Text = dr.GetString("password")
                txtbxfname.Text = dr.GetString("first_name")
                txtbxmname.Text = dr.GetString("middle_name")
                txtbxlname.Text = dr.GetString("last_name")
                txtbxadds.Text = dr.GetString("address")
                txtbxcontact.Text = dr.GetString("contact")
                txtbxemail.Text = dr.GetString("email")
                cmboSex.SelectedItem = dr.GetString("sex")
                txtbxeduc.Text = dr.GetString("educational")
                txtbxbdate.Value = dr.GetDateTime("birthdate")

                cmbxposition.SelectedItem = dr.GetString("position")
                cmbxemployment.SelectedItem = dr.GetString("employment_status")
                cmbxoffice.SelectedItem = dr.GetString("office")
                cmbxcomm.SelectedItem = dr.GetString("committee")
                cmbxcomm.SelectedItem = dr.GetString("committee")
            End While
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            conn.Close()
        End Try
    End Sub



    'SEARCH FUNCTION-------------------------------------------
    Private Sub search_TextChanged(sender As Object, e As EventArgs) Handles search.TextChanged
        BeneficiariesDGV.Rows.Clear()
        Try
            conn.Open() ' Opens a connection to the database.
            Dim cmd As New MySqlCommand("SELECT * FROM beneficiaries WHERE full_name LIKE '%" & search.Text & "%' AND user_id = @ID", conn) ' Creates a new MySqlCommand object to search for beneficiaries whose full_name matches the search text and belong to the current user (Form2.log_id).
            cmd.Parameters.AddWithValue("@ID", Form2.log_id) ' Passes the log_id from Form2 as a parameter to the query.
            dr = cmd.ExecuteReader ' Executes the query and populates the DataReader object "dr" with the results.
            While dr.Read ' While reading the results from the DataReader...
                BeneficiariesDGV.Rows.Add(dr.Item("user_id"), dr.Item("full_name"), dr.Item("age"), dr.Item("relationship")) ' Adds a new row to the BeneficiariesDGV DataGridView and populates it with the values from the current row of the DataReader.
            End While
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            conn.Close() ' Closes the database connection.
        End Try
    End Sub



    'UPDATE FUNCTION---------------------------------------------
    Public Sub Update()
        Dim location As String = My.Application.Info.DirectoryPath
        'Dim indext As Integer = locateProject.IndexOf("bin\Debug\net6.0-windows")
        'Dim location As String = locateProject.Substring(0, indext)
        Dim imageInput As String

        Dim random As New Random()
        Dim randomNum As Integer = random.Next(1, 501)
        Dim destinationPath As String = location & "\Others\images\" & txtbxusername.Text & randomNum & getExtension

        Try
            conn.Open() ' Opens a connection to the database.

            Dim selectCmd As New MySqlCommand("SELECT image from users WHERE id = @ID", conn)
            selectCmd.Parameters.AddWithValue("@ID", Form2.log_id)
            Dim dr As MySqlDataReader = selectCmd.ExecuteReader() ' Use a separate variable for the first DataReader
            Dim imageName As String = ""
            While dr.Read()
                imageName = dr.GetString("image")
            End While
            dr.Close() ' Close the first DataReader

            ' Check if the source file path is valid and exists
            If Not String.IsNullOrEmpty(sourceFilePath) AndAlso File.Exists(sourceFilePath) Then
                File.Copy(sourceFilePath, destinationPath, True)
                imageInput = "\" & txtbxusername.Text & randomNum & getExtension
            Else
                imageInput = imageName
            End If
            ' Create an update command to update the database
            Dim cmd As New MySqlCommand("UPDATE users " &
                                        "INNER JOIN user_info ON users.id = user_info.user_id " &
                                        "SET users.username = @username, users.password = @password, users.first_name = @first, users.middle_name = @mid, users.last_name = @last, users.position = @pos, users.image = @img, users.sex = @sex, user_info.address = @adds, user_info.contact = @contact, user_info.email = @email, user_info.educational = @educ, user_info.birthdate = @birthdate, user_info.office = @office, user_info.employment_status = @employ, user_info.committee = @comm " &
                                        "WHERE users.id = @ID", conn)
            cmd.Parameters.Clear()
            cmd.Parameters.AddWithValue("@ID", Form2.log_id)
            cmd.Parameters.AddWithValue("@username", txtbxusername.Text)
            cmd.Parameters.AddWithValue("@password", txtbxpassword.Text)
            cmd.Parameters.AddWithValue("@first", txtbxfname.Text)
            cmd.Parameters.AddWithValue("@mid", txtbxmname.Text)
            cmd.Parameters.AddWithValue("@last", txtbxlname.Text)
            cmd.Parameters.AddWithValue("@pos", cmbxposition.Text)
            cmd.Parameters.AddWithValue("@img", imageInput)

            cmd.Parameters.AddWithValue("@adds", txtbxadds.Text)
            cmd.Parameters.AddWithValue("@sex", cmboSex.SelectedItem)
            cmd.Parameters.AddWithValue("@contact", txtbxcontact.Text)
            cmd.Parameters.AddWithValue("@email", txtbxemail.Text)
            cmd.Parameters.AddWithValue("@educ", txtbxeduc.Text)
            cmd.Parameters.AddWithValue("@birthdate", txtbxbdate.Value)
            cmd.Parameters.AddWithValue("@office", cmbxoffice.SelectedItem)
            cmd.Parameters.AddWithValue("@employ", cmbxemployment.SelectedItem)
            cmd.Parameters.AddWithValue("@comm", cmbxcomm.SelectedItem)

            cmd.ExecuteNonQuery()  'Execute the update command
            MessageBox.Show("Updated successfully!", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            conn.Close() ' Closes the connection to the database.
        End Try

    End Sub

    Private Sub Panel2_Paint(sender As Object, e As PaintEventArgs) Handles Panel2.Paint
        Get_info()
    End Sub
    Private Sub Guna2Button3_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click
        Update()
    End Sub




    'EXPORT TO EXCEL-------------------------------------------------------------------------
    Public Sub SetEPPlusLicenseContext()
        ' Set the EPPlus license context to NonCommercial
        ' This allows the use of EPPlus library in a non-commercial context
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial
    End Sub
    Public Sub ExportToExcel(BenefeciariesDGV As DataGridView, filePath As String)
        ' Set the EPPlus license context
        SetEPPlusLicenseContext()

        ' Create a new Excel package
        Using package As New ExcelPackage()

            ' Add a worksheet to the package
            Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("Employees")

            ' Write column headers to the worksheet
            For j = 0 To BenefeciariesDGV.Columns.Count - 1
                worksheet.Cells(1, j + 1).Value = BenefeciariesDGV.Columns(j).HeaderText
            Next

            ' Write data rows to the worksheet
            For i = 0 To BenefeciariesDGV.Rows.Count - 1
                For j = 0 To BenefeciariesDGV.Columns.Count - 1
                    worksheet.Cells(i + 2, j + 1).Value = BenefeciariesDGV.Rows(i).Cells(j).Value
                Next
            Next

            ' Define a range for styling the header
            Dim headerRange As ExcelRange = worksheet.Cells(1, 1, 1, BenefeciariesDGV.Columns.Count)

            ' Set background color, font color, and border for the header
            headerRange.Style.Fill.PatternType = ExcelFillStyle.Solid
            headerRange.Style.Fill.BackgroundColor.SetColor(Color.LightGreen)
            headerRange.Style.Font.Color.SetColor(Color.Black)
            headerRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
            headerRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin

            ' Define a range for styling the data rows
            Dim dataRange As ExcelRange = worksheet.Cells(2, 1, BenefeciariesDGV.Rows.Count + 1, BenefeciariesDGV.Columns.Count)

            ' Set background color, font color, and border for the data rows
            dataRange.Style.Fill.PatternType = ExcelFillStyle.Solid
            dataRange.Style.Font.Color.SetColor(Color.Black)
            dataRange.Style.Fill.BackgroundColor.SetColor(Color.White)
            dataRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
            dataRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin

            ' Set column widths for specific columns
            worksheet.Column(1).Width = 7.43 ' Column A
            worksheet.Column(2).Width = 42 ' Column B
            worksheet.Column(3).Width = 8 ' Column C
            worksheet.Column(4).Width = 18.57 ' Column D

            ' Save the Excel package to the specified file path
            Dim fileInfo As New FileInfo(filePath)
            package.SaveAs(fileInfo)
        End Using

        ' Open the folder location of the exported Excel file
        Dim processStartInfo As New ProcessStartInfo()
        processStartInfo.FileName = "explorer.exe"
        processStartInfo.Arguments = "/select, """ & filePath & """"
        processStartInfo.UseShellExecute = True
        Process.Start(processStartInfo)
    End Sub

    'EXPORT TO EXCEL BUTTON--------------------------
    Private Sub Guna2Button2_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        Try
            conn.Open() ' Opens a connection to the database.
            Dim cmd As New MySqlCommand("SELECT * FROM users WHERE id=@id", conn) ' Creates a new MySqlCommand object to retrieve data from the "users" table where the id matches Form2.log_id.
            cmd.Parameters.AddWithValue("@id", Form2.log_id) ' Passes the log_id from Form2 as a parameter to the query.
            dr = cmd.ExecuteReader ' Executes the query and populates the DataReader object "dr" with the results.
            While dr.Read ' While reading the results from the DataReader...
                Dim documentsPath As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) ' Retrieves the path for the My Documents folder.
                Dim filePath As String = Path.Combine(documentsPath, dr.GetString("first_name") & "_beneficiary.xlsx") ' Combines the My Documents folder path with the user's first name from the query results. This will be the location and name of the output Excel file.
                Dim fileName As String = dr.GetString("first_name") & "_beneficiary.xlsx" ' The name of the output Excel file based on the user's first name from the query results.
                If File.Exists(filePath) Then ' IF THE FILE ALREADY EXISTS -------
                    ' If yes, displays an alert box indicating that the file already exists along with its location.
                    MessageBox.Show("The file already exists; this is the file location" & vbCrLf & filePath, "Information", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Else
                    ExportToExcel(BeneficiariesDGV, filePath) ' If the file doesn't exist yet, performs the export to Excel using the ExportToExcel function. Passes BeneficiariesDGV (DataGridView) and the path of the output file as arguments.
                    MessageBox.Show("Export completed. The file name is " & fileName, "Excel file", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    ' Displays an alert box indicating that the export is completed and also shows the name of the output file.
                End If
            End While
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error) ' Closes the database connection.
        Finally
            conn.Close()
        End Try
    End Sub

    '---------NUMBER ONLY AND LETTER ONLY------------'
    Private Sub txtbxcontact_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtbxcontact.KeyPress
        If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub
    Private Shared Sub txtbxfname_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtbxfname.KeyPress
        If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsLetter(e.KeyChar) AndAlso Not Char.IsWhiteSpace(e.KeyChar) AndAlso Not Char.IsPunctuation(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub
    Private Sub txtbxmname_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtbxmname.KeyPress
        txtbxfname_KeyPress(sender, e)
    End Sub
    Private Sub txtbxlname_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtbxlname.KeyPress
        txtbxfname_KeyPress(sender, e)
    End Sub


    ''Email Edit Validation
    Private Function IsValidEmail(email As String) As Boolean
        ' Create a regular expression pattern to validate email addresses
        Dim emailRegex As New Regex("^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$")

        ' Check if the email matches the regular expression pattern
        Return emailRegex.IsMatch(email)
    End Function

    Private Sub txtEmail_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtbxemail.Validating
        ' Get the trimmed input email from the textbox
        Dim inputEmail As String = txtbxemail.Text.Trim()

        ' Check if the textbox is empty
        If txtbxemail.Text = "" Then
            txtbxemail.Text = txtbxemail.Text
        Else
            ' Check if the input email is a valid email address using the IsValidEmail function
            If Not IsValidEmail(inputEmail) Then
                ' Show a message box indicating an invalid email address
                MessageBox.Show("Invalid email address." & vbCrLf & "Please enter a valid email address.", "INFORMATION", MessageBoxButtons.OK, MessageBoxIcon.Hand)

                ' Cancel the event to prevent losing focus from the textbox
                e.Cancel = True
            End If
        End If
    End Sub


    Private Sub Guna2Button4_Click(sender As Object, e As EventArgs) Handles Guna2Button4.Click
        ' Create a new instance of the OpenFileDialog class
        Dim opf As New OpenFileDialog

        ' Set the filter to restrict the file types that can be selected
        opf.Filter = "Choose Image(*.jpg; *.png; *.gif) | * .jpg; *.png; *.gif"

        ' Display the OpenFileDialog and check if the user clicked "OK"
        If opf.ShowDialog = DialogResult.OK Then
            ' Retrieve the full path of the selected file
            sourceFilePath = System.IO.Path.GetFullPath(opf.FileName)

            ' Set the BackgroundImage property of a control to the selected image
            user_Profile.BackgroundImage = Image.FromFile(sourceFilePath)

            ' Retrieve the file extension of the selected file
            getExtension = System.IO.Path.GetExtension(opf.FileName)
        End If
    End Sub

    Private Sub Guna2ControlBox1_Click(sender As Object, e As EventArgs) Handles Guna2ControlBox1.Click
        Application.Exit()
    End Sub
End Class
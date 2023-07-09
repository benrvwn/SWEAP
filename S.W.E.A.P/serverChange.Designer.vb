<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class serverChange
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        components = New ComponentModel.Container()
        Dim resources As ComponentModel.ComponentResourceManager = New ComponentModel.ComponentResourceManager(GetType(serverChange))
        Dim CustomizableEdges13 As Guna.UI2.WinForms.Suite.CustomizableEdges = New Guna.UI2.WinForms.Suite.CustomizableEdges()
        Dim CustomizableEdges14 As Guna.UI2.WinForms.Suite.CustomizableEdges = New Guna.UI2.WinForms.Suite.CustomizableEdges()
        Dim CustomizableEdges11 As Guna.UI2.WinForms.Suite.CustomizableEdges = New Guna.UI2.WinForms.Suite.CustomizableEdges()
        Dim CustomizableEdges12 As Guna.UI2.WinForms.Suite.CustomizableEdges = New Guna.UI2.WinForms.Suite.CustomizableEdges()
        Dim CustomizableEdges1 As Guna.UI2.WinForms.Suite.CustomizableEdges = New Guna.UI2.WinForms.Suite.CustomizableEdges()
        Dim CustomizableEdges2 As Guna.UI2.WinForms.Suite.CustomizableEdges = New Guna.UI2.WinForms.Suite.CustomizableEdges()
        Dim CustomizableEdges3 As Guna.UI2.WinForms.Suite.CustomizableEdges = New Guna.UI2.WinForms.Suite.CustomizableEdges()
        Dim CustomizableEdges4 As Guna.UI2.WinForms.Suite.CustomizableEdges = New Guna.UI2.WinForms.Suite.CustomizableEdges()
        Dim CustomizableEdges5 As Guna.UI2.WinForms.Suite.CustomizableEdges = New Guna.UI2.WinForms.Suite.CustomizableEdges()
        Dim CustomizableEdges6 As Guna.UI2.WinForms.Suite.CustomizableEdges = New Guna.UI2.WinForms.Suite.CustomizableEdges()
        Dim CustomizableEdges7 As Guna.UI2.WinForms.Suite.CustomizableEdges = New Guna.UI2.WinForms.Suite.CustomizableEdges()
        Dim CustomizableEdges8 As Guna.UI2.WinForms.Suite.CustomizableEdges = New Guna.UI2.WinForms.Suite.CustomizableEdges()
        Dim CustomizableEdges9 As Guna.UI2.WinForms.Suite.CustomizableEdges = New Guna.UI2.WinForms.Suite.CustomizableEdges()
        Dim CustomizableEdges10 As Guna.UI2.WinForms.Suite.CustomizableEdges = New Guna.UI2.WinForms.Suite.CustomizableEdges()
        Guna2Elipse1 = New Guna.UI2.WinForms.Guna2Elipse(components)
        Label1 = New Label()
        Label2 = New Label()
        Guna2Button1 = New Guna.UI2.WinForms.Guna2Button()
        Guna2Button2 = New Guna.UI2.WinForms.Guna2Button()
        Panel1 = New Panel()
        Label6 = New Label()
        Label5 = New Label()
        Label4 = New Label()
        Label3 = New Label()
        IP = New Label()
        txtDatabase = New Guna.UI2.WinForms.Guna2TextBox()
        txtPassword = New Guna.UI2.WinForms.Guna2TextBox()
        txtUsername = New Guna.UI2.WinForms.Guna2TextBox()
        txtPort = New Guna.UI2.WinForms.Guna2TextBox()
        txtIp = New Guna.UI2.WinForms.Guna2TextBox()
        Panel1.SuspendLayout()
        SuspendLayout()
        ' 
        ' Guna2Elipse1
        ' 
        Guna2Elipse1.TargetControl = Me
        ' 
        ' Label1
        ' 
        Label1.Dock = DockStyle.Top
        Label1.Font = New Font("Segoe UI", 14F, FontStyle.Regular, GraphicsUnit.Point)
        Label1.ForeColor = SystemColors.ActiveCaptionText
        Label1.Location = New Point(0, 0)
        Label1.Name = "Label1"
        Label1.Size = New Size(372, 29)
        Label1.TabIndex = 0
        Label1.Text = "Service and port settings"
        Label1.TextAlign = ContentAlignment.MiddleLeft
        ' 
        ' Label2
        ' 
        Label2.AutoSize = True
        Label2.Font = New Font("Segoe UI", 8F, FontStyle.Regular, GraphicsUnit.Point)
        Label2.ForeColor = SystemColors.ActiveCaptionText
        Label2.Location = New Point(4, 47)
        Label2.Name = "Label2"
        Label2.Size = New Size(367, 91)
        Label2.TabIndex = 2
        Label2.Text = resources.GetString("Label2.Text")
        Label2.TextAlign = ContentAlignment.MiddleCenter
        ' 
        ' Guna2Button1
        ' 
        Guna2Button1.BorderThickness = 2
        Guna2Button1.CustomizableEdges = CustomizableEdges13
        Guna2Button1.DisabledState.BorderColor = Color.DarkGray
        Guna2Button1.DisabledState.CustomBorderColor = Color.DarkGray
        Guna2Button1.DisabledState.FillColor = Color.FromArgb(CByte(169), CByte(169), CByte(169))
        Guna2Button1.DisabledState.ForeColor = Color.FromArgb(CByte(141), CByte(141), CByte(141))
        Guna2Button1.FillColor = SystemColors.ControlLight
        Guna2Button1.Font = New Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point)
        Guna2Button1.ForeColor = Color.Black
        Guna2Button1.Image = My.Resources.Resources.check1
        Guna2Button1.Location = New Point(251, 471)
        Guna2Button1.Name = "Guna2Button1"
        Guna2Button1.ShadowDecoration.CustomizableEdges = CustomizableEdges14
        Guna2Button1.Size = New Size(109, 23)
        Guna2Button1.TabIndex = 3
        Guna2Button1.Text = "Save"
        ' 
        ' Guna2Button2
        ' 
        Guna2Button2.BorderThickness = 2
        Guna2Button2.CustomizableEdges = CustomizableEdges11
        Guna2Button2.DisabledState.BorderColor = Color.DarkGray
        Guna2Button2.DisabledState.CustomBorderColor = Color.DarkGray
        Guna2Button2.DisabledState.FillColor = Color.FromArgb(CByte(169), CByte(169), CByte(169))
        Guna2Button2.DisabledState.ForeColor = Color.FromArgb(CByte(141), CByte(141), CByte(141))
        Guna2Button2.FillColor = SystemColors.ControlLight
        Guna2Button2.Font = New Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point)
        Guna2Button2.ForeColor = Color.Black
        Guna2Button2.Image = My.Resources.Resources.cross_button1
        Guna2Button2.Location = New Point(136, 471)
        Guna2Button2.Name = "Guna2Button2"
        Guna2Button2.ShadowDecoration.CustomizableEdges = CustomizableEdges12
        Guna2Button2.Size = New Size(109, 23)
        Guna2Button2.TabIndex = 4
        Guna2Button2.Text = "Abort"
        ' 
        ' Panel1
        ' 
        Panel1.BackColor = Color.FromArgb(CByte(54), CByte(69), CByte(94))
        Panel1.Controls.Add(Label6)
        Panel1.Controls.Add(Label5)
        Panel1.Controls.Add(Label4)
        Panel1.Controls.Add(Label3)
        Panel1.Controls.Add(IP)
        Panel1.Controls.Add(txtDatabase)
        Panel1.Controls.Add(txtPassword)
        Panel1.Controls.Add(txtUsername)
        Panel1.Controls.Add(txtPort)
        Panel1.Controls.Add(txtIp)
        Panel1.Location = New Point(0, 150)
        Panel1.Name = "Panel1"
        Panel1.Size = New Size(372, 306)
        Panel1.TabIndex = 5
        ' 
        ' Label6
        ' 
        Label6.AutoSize = True
        Label6.Font = New Font("Segoe UI", 12F, FontStyle.Regular, GraphicsUnit.Point)
        Label6.ForeColor = SystemColors.ControlLightLight
        Label6.Location = New Point(22, 243)
        Label6.Name = "Label6"
        Label6.Size = New Size(77, 21)
        Label6.TabIndex = 1
        Label6.Text = "Database:"
        ' 
        ' Label5
        ' 
        Label5.AutoSize = True
        Label5.Font = New Font("Segoe UI", 12F, FontStyle.Regular, GraphicsUnit.Point)
        Label5.ForeColor = SystemColors.ControlLightLight
        Label5.Location = New Point(22, 191)
        Label5.Name = "Label5"
        Label5.Size = New Size(79, 21)
        Label5.TabIndex = 1
        Label5.Text = "Password:"
        ' 
        ' Label4
        ' 
        Label4.AutoSize = True
        Label4.Font = New Font("Segoe UI", 12F, FontStyle.Regular, GraphicsUnit.Point)
        Label4.ForeColor = SystemColors.ControlLightLight
        Label4.Location = New Point(22, 139)
        Label4.Name = "Label4"
        Label4.Size = New Size(84, 21)
        Label4.TabIndex = 1
        Label4.Text = "Username:"
        ' 
        ' Label3
        ' 
        Label3.AutoSize = True
        Label3.Font = New Font("Segoe UI", 12F, FontStyle.Regular, GraphicsUnit.Point)
        Label3.ForeColor = SystemColors.ControlLightLight
        Label3.Location = New Point(22, 88)
        Label3.Name = "Label3"
        Label3.Size = New Size(41, 21)
        Label3.TabIndex = 1
        Label3.Text = "Port:"
        ' 
        ' IP
        ' 
        IP.AutoSize = True
        IP.Font = New Font("Segoe UI", 12F, FontStyle.Regular, GraphicsUnit.Point)
        IP.ForeColor = SystemColors.ControlLightLight
        IP.Location = New Point(22, 37)
        IP.Name = "IP"
        IP.Size = New Size(26, 21)
        IP.TabIndex = 1
        IP.Text = "IP:"
        ' 
        ' txtDatabase
        ' 
        txtDatabase.CustomizableEdges = CustomizableEdges1
        txtDatabase.DefaultText = ""
        txtDatabase.DisabledState.BorderColor = Color.FromArgb(CByte(208), CByte(208), CByte(208))
        txtDatabase.DisabledState.FillColor = Color.FromArgb(CByte(226), CByte(226), CByte(226))
        txtDatabase.DisabledState.ForeColor = Color.FromArgb(CByte(138), CByte(138), CByte(138))
        txtDatabase.DisabledState.PlaceholderForeColor = Color.FromArgb(CByte(138), CByte(138), CByte(138))
        txtDatabase.FocusedState.BorderColor = Color.FromArgb(CByte(94), CByte(148), CByte(255))
        txtDatabase.Font = New Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point)
        txtDatabase.HoverState.BorderColor = Color.FromArgb(CByte(94), CByte(148), CByte(255))
        txtDatabase.Location = New Point(136, 236)
        txtDatabase.Name = "txtDatabase"
        txtDatabase.PasswordChar = ChrW(0)
        txtDatabase.PlaceholderText = ""
        txtDatabase.SelectedText = ""
        txtDatabase.ShadowDecoration.CustomizableEdges = CustomizableEdges2
        txtDatabase.Size = New Size(211, 36)
        txtDatabase.TabIndex = 0
        ' 
        ' txtPassword
        ' 
        txtPassword.CustomizableEdges = CustomizableEdges3
        txtPassword.DefaultText = ""
        txtPassword.DisabledState.BorderColor = Color.FromArgb(CByte(208), CByte(208), CByte(208))
        txtPassword.DisabledState.FillColor = Color.FromArgb(CByte(226), CByte(226), CByte(226))
        txtPassword.DisabledState.ForeColor = Color.FromArgb(CByte(138), CByte(138), CByte(138))
        txtPassword.DisabledState.PlaceholderForeColor = Color.FromArgb(CByte(138), CByte(138), CByte(138))
        txtPassword.FocusedState.BorderColor = Color.FromArgb(CByte(94), CByte(148), CByte(255))
        txtPassword.Font = New Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point)
        txtPassword.HoverState.BorderColor = Color.FromArgb(CByte(94), CByte(148), CByte(255))
        txtPassword.Location = New Point(136, 184)
        txtPassword.Name = "txtPassword"
        txtPassword.PasswordChar = ChrW(0)
        txtPassword.PlaceholderText = ""
        txtPassword.SelectedText = ""
        txtPassword.ShadowDecoration.CustomizableEdges = CustomizableEdges4
        txtPassword.Size = New Size(211, 36)
        txtPassword.TabIndex = 0
        ' 
        ' txtUsername
        ' 
        txtUsername.CustomizableEdges = CustomizableEdges5
        txtUsername.DefaultText = ""
        txtUsername.DisabledState.BorderColor = Color.FromArgb(CByte(208), CByte(208), CByte(208))
        txtUsername.DisabledState.FillColor = Color.FromArgb(CByte(226), CByte(226), CByte(226))
        txtUsername.DisabledState.ForeColor = Color.FromArgb(CByte(138), CByte(138), CByte(138))
        txtUsername.DisabledState.PlaceholderForeColor = Color.FromArgb(CByte(138), CByte(138), CByte(138))
        txtUsername.FocusedState.BorderColor = Color.FromArgb(CByte(94), CByte(148), CByte(255))
        txtUsername.Font = New Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point)
        txtUsername.HoverState.BorderColor = Color.FromArgb(CByte(94), CByte(148), CByte(255))
        txtUsername.Location = New Point(136, 132)
        txtUsername.Name = "txtUsername"
        txtUsername.PasswordChar = ChrW(0)
        txtUsername.PlaceholderText = ""
        txtUsername.SelectedText = ""
        txtUsername.ShadowDecoration.CustomizableEdges = CustomizableEdges6
        txtUsername.Size = New Size(211, 36)
        txtUsername.TabIndex = 0
        ' 
        ' txtPort
        ' 
        txtPort.CustomizableEdges = CustomizableEdges7
        txtPort.DefaultText = ""
        txtPort.DisabledState.BorderColor = Color.FromArgb(CByte(208), CByte(208), CByte(208))
        txtPort.DisabledState.FillColor = Color.FromArgb(CByte(226), CByte(226), CByte(226))
        txtPort.DisabledState.ForeColor = Color.FromArgb(CByte(138), CByte(138), CByte(138))
        txtPort.DisabledState.PlaceholderForeColor = Color.FromArgb(CByte(138), CByte(138), CByte(138))
        txtPort.FocusedState.BorderColor = Color.FromArgb(CByte(94), CByte(148), CByte(255))
        txtPort.Font = New Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point)
        txtPort.HoverState.BorderColor = Color.FromArgb(CByte(94), CByte(148), CByte(255))
        txtPort.Location = New Point(136, 81)
        txtPort.Name = "txtPort"
        txtPort.PasswordChar = ChrW(0)
        txtPort.PlaceholderText = ""
        txtPort.SelectedText = ""
        txtPort.ShadowDecoration.CustomizableEdges = CustomizableEdges8
        txtPort.Size = New Size(211, 36)
        txtPort.TabIndex = 0
        ' 
        ' txtIp
        ' 
        txtIp.CustomizableEdges = CustomizableEdges9
        txtIp.DefaultText = ""
        txtIp.DisabledState.BorderColor = Color.FromArgb(CByte(208), CByte(208), CByte(208))
        txtIp.DisabledState.FillColor = Color.FromArgb(CByte(226), CByte(226), CByte(226))
        txtIp.DisabledState.ForeColor = Color.FromArgb(CByte(138), CByte(138), CByte(138))
        txtIp.DisabledState.PlaceholderForeColor = Color.FromArgb(CByte(138), CByte(138), CByte(138))
        txtIp.FocusedState.BorderColor = Color.FromArgb(CByte(94), CByte(148), CByte(255))
        txtIp.Font = New Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point)
        txtIp.HoverState.BorderColor = Color.FromArgb(CByte(94), CByte(148), CByte(255))
        txtIp.Location = New Point(136, 30)
        txtIp.Name = "txtIp"
        txtIp.PasswordChar = ChrW(0)
        txtIp.PlaceholderText = ""
        txtIp.SelectedText = ""
        txtIp.ShadowDecoration.CustomizableEdges = CustomizableEdges10
        txtIp.Size = New Size(211, 36)
        txtIp.TabIndex = 0
        ' 
        ' serverChange
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        BackColor = Color.FromArgb(CByte(224), CByte(224), CByte(224))
        ClientSize = New Size(372, 510)
        Controls.Add(Panel1)
        Controls.Add(Guna2Button2)
        Controls.Add(Guna2Button1)
        Controls.Add(Label2)
        Controls.Add(Label1)
        FormBorderStyle = FormBorderStyle.None
        Name = "serverChange"
        StartPosition = FormStartPosition.CenterScreen
        Text = "serverChange"
        Panel1.ResumeLayout(False)
        Panel1.PerformLayout()
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents Guna2Elipse1 As Guna.UI2.WinForms.Guna2Elipse
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Guna2Button1 As Guna.UI2.WinForms.Guna2Button
    Friend WithEvents Guna2Button2 As Guna.UI2.WinForms.Guna2Button
    Friend WithEvents Panel1 As Panel
    Friend WithEvents txtIp As Guna.UI2.WinForms.Guna2TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents IP As Label
    Friend WithEvents txtDatabase As Guna.UI2.WinForms.Guna2TextBox
    Friend WithEvents txtPassword As Guna.UI2.WinForms.Guna2TextBox
    Friend WithEvents txtUsername As Guna.UI2.WinForms.Guna2TextBox
    Friend WithEvents txtPort As Guna.UI2.WinForms.Guna2TextBox
End Class

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMain
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
        Me.TxtLongitude = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TxtLatitude = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.BtnCalculate = New System.Windows.Forms.Button()
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker()
        Me.TxtSunrise = New System.Windows.Forms.TextBox()
        Me.TxtSunset = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.CmbCity = New System.Windows.Forms.ComboBox()
        Me.LblCity = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.LblTimeZone = New System.Windows.Forms.Label()
        Me.TxtHighNoon = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.TxtTimezone = New System.Windows.Forms.TextBox()
        Me.fSaveFile = New System.Windows.Forms.SaveFileDialog()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'TxtLongitude
        '
        Me.TxtLongitude.Location = New System.Drawing.Point(83, 81)
        Me.TxtLongitude.Name = "TxtLongitude"
        Me.TxtLongitude.Size = New System.Drawing.Size(118, 20)
        Me.TxtLongitude.TabIndex = 0
        Me.TxtLongitude.UseWaitCursor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(83, 65)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(54, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Longitude"
        '
        'TxtLatitude
        '
        Me.TxtLatitude.Location = New System.Drawing.Point(232, 81)
        Me.TxtLatitude.Name = "TxtLatitude"
        Me.TxtLatitude.Size = New System.Drawing.Size(114, 20)
        Me.TxtLatitude.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(232, 65)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(45, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Latitude"
        '
        'BtnCalculate
        '
        Me.BtnCalculate.FlatAppearance.BorderColor = System.Drawing.Color.Maroon
        Me.BtnCalculate.FlatAppearance.BorderSize = 2
        Me.BtnCalculate.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.BtnCalculate.Location = New System.Drawing.Point(391, 80)
        Me.BtnCalculate.Name = "BtnCalculate"
        Me.BtnCalculate.Size = New System.Drawing.Size(118, 23)
        Me.BtnCalculate.TabIndex = 4
        Me.BtnCalculate.Text = "Calculate"
        Me.BtnCalculate.UseVisualStyleBackColor = True
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.AllowDrop = True
        Me.DateTimePicker1.CustomFormat = "M'/'dd'/'yyyy hh':'mm tt"
        Me.DateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DateTimePicker1.Location = New System.Drawing.Point(83, 136)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.Size = New System.Drawing.Size(118, 20)
        Me.DateTimePicker1.TabIndex = 5
        '
        'TxtSunrise
        '
        Me.TxtSunrise.BackColor = System.Drawing.SystemColors.Control
        Me.TxtSunrise.Location = New System.Drawing.Point(573, 38)
        Me.TxtSunrise.Name = "TxtSunrise"
        Me.TxtSunrise.Size = New System.Drawing.Size(191, 20)
        Me.TxtSunrise.TabIndex = 6
        '
        'TxtSunset
        '
        Me.TxtSunset.Location = New System.Drawing.Point(573, 83)
        Me.TxtSunset.Name = "TxtSunset"
        Me.TxtSunset.ReadOnly = True
        Me.TxtSunset.Size = New System.Drawing.Size(191, 20)
        Me.TxtSunset.TabIndex = 7
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(573, 21)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(42, 13)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "Sunrise"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(573, 67)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(40, 13)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "Sunset"
        '
        'CmbCity
        '
        Me.CmbCity.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbCity.FormattingEnabled = True
        Me.CmbCity.Location = New System.Drawing.Point(83, 37)
        Me.CmbCity.Name = "CmbCity"
        Me.CmbCity.Size = New System.Drawing.Size(118, 21)
        Me.CmbCity.TabIndex = 12
        '
        'LblCity
        '
        Me.LblCity.AutoSize = True
        Me.LblCity.Location = New System.Drawing.Point(83, 21)
        Me.LblCity.Name = "LblCity"
        Me.LblCity.Size = New System.Drawing.Size(24, 13)
        Me.LblCity.TabIndex = 13
        Me.LblCity.Text = "City"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(83, 120)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(30, 13)
        Me.Label5.TabIndex = 14
        Me.Label5.Text = "Date"
        '
        'LblTimeZone
        '
        Me.LblTimeZone.AutoSize = True
        Me.LblTimeZone.Location = New System.Drawing.Point(229, 123)
        Me.LblTimeZone.Name = "LblTimeZone"
        Me.LblTimeZone.Size = New System.Drawing.Size(58, 13)
        Me.LblTimeZone.TabIndex = 15
        Me.LblTimeZone.Text = "Time Zone"
        '
        'TxtHighNoon
        '
        Me.TxtHighNoon.Location = New System.Drawing.Point(573, 136)
        Me.TxtHighNoon.Name = "TxtHighNoon"
        Me.TxtHighNoon.ReadOnly = True
        Me.TxtHighNoon.Size = New System.Drawing.Size(191, 20)
        Me.TxtHighNoon.TabIndex = 16
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(573, 120)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(58, 13)
        Me.Label6.TabIndex = 17
        Me.Label6.Text = "High Noon"
        '
        'TxtTimezone
        '
        Me.TxtTimezone.Location = New System.Drawing.Point(231, 136)
        Me.TxtTimezone.Name = "TxtTimezone"
        Me.TxtTimezone.ReadOnly = True
        Me.TxtTimezone.Size = New System.Drawing.Size(115, 20)
        Me.TxtTimezone.TabIndex = 18
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(86, 214)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(260, 23)
        Me.Button1.TabIndex = 19
        Me.Button1.Text = "Generate one year file"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'FrmMain
        '
        Me.AllowDrop = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.TxtTimezone)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.TxtHighNoon)
        Me.Controls.Add(Me.LblTimeZone)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.LblCity)
        Me.Controls.Add(Me.CmbCity)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.TxtSunset)
        Me.Controls.Add(Me.TxtSunrise)
        Me.Controls.Add(Me.DateTimePicker1)
        Me.Controls.Add(Me.BtnCalculate)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TxtLatitude)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TxtLongitude)
        Me.Name = "FrmMain"
        Me.Text = "Sun Position Calculator"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents TxtLongitude As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents TxtLatitude As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents BtnCalculate As Button
    Friend WithEvents DateTimePicker1 As DateTimePicker
    Friend WithEvents TxtSunrise As TextBox
    Friend WithEvents TxtSunset As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents CmbCity As ComboBox
    Friend WithEvents LblCity As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents LblTimeZone As Label
    Friend WithEvents TxtHighNoon As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents TxtTimezone As TextBox
    Friend WithEvents fSaveFile As SaveFileDialog
    Friend WithEvents Button1 As Button
End Class

#region prereqs
[reflection.assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null

#MainPage
$MainMenu = New-Object System.Windows.Forms.Form
$MainMenu.Text = 'Special ID Creation'

$MainMenu.Size = '420,350'


#Label
$Label = New-Object System.Windows.Forms.Label
$Label.Location = '50,10'
$Label.Size = '300,30'
$Label.ForeColor = 'orange'
$Label.BackColor = 'black'
$Label.Text = "Please Choose the Type of Special ID"

$MainMenu.Controls.Add($Label)

$SidButton = New-Object System.Windows.Forms.Button
$SidButton.Location = '50,50'
$SidButton.size = '300,30'
$SidButton.Text = "Service ID(SID) Creation"

$MainMenu.Controls.Add($SidButton)

$GidButton = New-Object System.Windows.Forms.Button
$GidButton.Location = '50,90'
$GidButton.size = '300,30'
$GidButton.Text = "GROUP ID(GID) Creation"

$MainMenu.Controls.Add($GidButton)

$FTPButton = New-Object System.Windows.Forms.Button
$FTPButton.Location = '50,130'
$FTPButton.size = '300,30'
$FTPButton.Text = "FTP ID Creation"

$MainMenu.Controls.Add($FTPButton)

$KioskButton = New-Object System.Windows.Forms.Button
$KioskButton.Location = '50,170'
$KioskButton.size = '300,30'
$KioskButton.Text = "KIOSK ID Creation"

$MainMenu.Controls.Add($KioskButton)

$HelpButton = New-Object System.Windows.Forms.Button
$HelpButton.Location = '50,210'
$HelpButton.size = '300,30'
$HelpButton.Text = "Help/Suggestion/Feedback"

$MainMenu.Controls.Add($HelpButton)

$ExitButton = New-Object System.Windows.Forms.Button
$ExitButton.Location = '50,250'
$ExitButton.size = '300,30'
$ExitButton.Text = "Exit Menu"

$MainMenu.Controls.Add($ExitButton)
$SID = "C:\Special_Id_creationV2.0\SIDMenu.ps1"
$GID = "C:\Special_Id_creationV2.0\GIDMenu.ps1"
$KIOSK = "C:\Special_Id_creationV2.0\KIOSKMenu.ps1"
$FTP = "C:\Special_Id_creationV2.0\FTPMenu.ps1"
$Help = "C:\Special_Id_creationV2.0\Help_suggestion.Ps1"




#Adding Logic

$SidButton.Add_click({
$MainMenu.Hide()
Invoke-Expression $SID
})

$GidButton.Add_click({
$MainMenu.Hide()
Invoke-Expression $GID
})

$KioskButton.Add_click({
$MainMenu.Hide()
Invoke-Expression $KIOSK
})

$FTPButton.Add_click({
$MainMenu.Hide()
Invoke-Expression $FTP
})

$HelpButton.Add_click({
$MainMenu.Hide()
Invoke-Expression $Help
})

$ExitButton.Add_Click({
$MainMenu.Close()
})


$MainMenu.ShowDialog()
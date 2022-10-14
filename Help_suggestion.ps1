# Menu

[reflection.assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
Add-Type -AssemblyName PresentationFramework

$HelpMenu = New-Object System.Windows.Forms.form
$HelpMenu.Text = "Feedback Form"
$HelpMenu.Size = '800,450'

$Help_Label = New-Object System.Windows.Forms.Label
$Help_Label.Text = "Please Enter Your Query/Suggestion/Feedback Below:"
$Help_Label.AutoSize = $True
$Help_Label.Location = '50,30'

$HelpMenu.Controls.Add($Help_Label)

$Help_text = New-Object System.Windows.Forms.TextBox
$Help_text.Size = New-Object System.Drawing.Size(600,250)
$Help_text.Multiline = $True
$Help_text.Location = '50,70'

$HelpMenu.Controls.Add($Help_text)

$cancelbutton = New-Object System.Windows.Forms.Button
$cancelbutton.AutoSize = $true
$cancelbutton.Text = 'cancel'
$cancelbutton.Location = '575,350'
$HelpMenu.Controls.Add($cancelbutton)

$SendMailButton = New-Object System.Windows.Forms.Button
$SendMailButton.Text = "Send Email"
$SendMailButton.Location = '425,350'
$SendMailButton.AutoSize = $true

$HelpMenu.Controls.Add($SendMailButton)

$HelpMenu.CancelButton = $cancelbutton

#Send E-Mail Function

Function Send_Email{
$smtpServer = "smtp.aptiv.com"
$smtpFrom = "Special_ID_Feedback@aptiv.com"
$smtpTo = $to
$smtpbcc = $bcc
$messageSubject = $subject
$messageBody = $body

$SMTPMessage = New-Object System.Net.Mail.MailMessage $smtpFrom, $smtpTo, $messageSubject, $messageBody
$SMTPMessage.bcc.Add($bcc)
 
$smtp = New-Object Net.Mail.SmtpClient($smtpServer)
$smtp.Send($SMTPMessage)

}
$SendMailButton.add_click({

$to = "karan.apoorva@aptiv.com"
$bcc = "laeequeahamed.basha@hcl.com"
$Subject = "Special ID Feedback"
$body = $Help_text.Text

Send_Email

[void][System.Windows.MessageBox]::Show( "Mail Sent Successfully ", "Feedback E-Mail", "OK", "Information" )

})

$HelpMenu.ShowDialog()
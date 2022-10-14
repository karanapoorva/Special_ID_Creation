# Sid Creation Menu
$filename = "Log_SID-"+ (Get-Date -Format "MM-dd-yyyy_HHmmss")
$logfile = "C:\Special_Id_creationV2.0\Logs\"+$filename+".log"

Start-Transcript -Path $logfile -Append

#GUI Part

#region prereqs
[reflection.assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
Add-Type -AssemblyName PresentationFramework

$SidMenu = New-Object System.Windows.Forms.form
$SidMenu.Text = "Service ID(SID) Creation"
$SidMenu.Size = '800,450'

# Label Box

$NetidLabel = New-Object System.Windows.Forms.Label
$NetidLabel.Location = '50,30'
$NetidLabel.Size = '250,30'
$NetidLabel.Text = "Enter Net id/ Logon Name :"

$SidMenu.Controls.Add($NetidLabel)

$DescriptionLabel = New-Object System.Windows.Forms.Label
$DescriptionLabel.Location = '50,80'
$DescriptionLabel.Size = '250,30'
$DescriptionLabel.Text = "Enter in the User Description :"

$SidMenu.Controls.Add($DescriptionLabel)

$TaskLabel = New-Object System.Windows.Forms.Label
$TaskLabel.Location = '50,130'
$TaskLabel.Size = '250,30'
$TaskLabel.Text = "Enter Task Number :"

$SidMenu.Controls.Add($TaskLabel)

$WorkstationLabel = New-Object System.Windows.Forms.Label
$WorkstationLabel.Location = '50,180'
$WorkstationLabel.Size = '300,30'
$WorkstationLabel.Text = "List of machines separated with (,) :"

$SidMenu.Controls.Add($WorkstationLabel)

$OwnerLabel = New-Object System.Windows.Forms.Label
$OwnerLabel.Location = '50,230'
$OwnerLabel.Size = '300,30'
$OwnerLabel.Text = "Enter Owner's Net Id :"

$SidMenu.Controls.Add($OwnerLabel)

$PswdLabel = New-Object System.Windows.Forms.Label
$PswdLabel.Location = '50,280'
$PswdLabel.Size = '300,30'
$PswdLabel.Text = "Enter any 8 Char complex Password :"

$SidMenu.Controls.Add($PswdLabel)

#Text Box

$netidtxtbox = New-Object System.Windows.Forms.TextBox
$netidtxtbox.Location = '400,30'
$netidtxtbox.Size = '250,30'

$SidMenu.Controls.Add($netidtxtbox)

$validateuserlabel = New-Object System.Windows.Forms.Label
$validateuserlabel.Location = '400,50'
$validateuserlabel.AutoSize = $true

$SidMenu.Controls.Add($validateuserlabel)

$Desctxtbox = New-Object System.Windows.Forms.TextBox
$Desctxtbox.Location = '400,80'
$Desctxtbox.Size = '250,30'

$SidMenu.Controls.Add($Desctxtbox)

$Tsktxtbox = New-Object System.Windows.Forms.TextBox
$Tsktxtbox.Location = '400,130'
$Tsktxtbox.Size = '250,30'

$SidMenu.Controls.Add($Tsktxtbox)

$Wstntxtbox = New-Object System.Windows.Forms.TextBox
$Wstntxtbox.Location = '400,180'
$Wstntxtbox.Size = '250,30'

$SidMenu.Controls.Add($Wstntxtbox)

$Ownertxtbox = New-Object System.Windows.Forms.TextBox
$Ownertxtbox.Location = '400,230'
$Ownertxtbox.Size = '250,30'

$SidMenu.Controls.Add($Ownertxtbox)

$ValidateOwner = New-Object System.Windows.Forms.Label
$ValidateOwner.Location = '400,250'
$ValidateOwner.AutoSize = $true

$SidMenu.Controls.Add($ValidateOwner)

$Pswdtxtbox = New-Object System.Windows.Forms.TextBox
$Pswdtxtbox.Location = '400,280'
$Pswdtxtbox.Size = '250,30'

$SidMenu.Controls.Add($Pswdtxtbox)

$PswdValLabel = New-Object System.Windows.Forms.Label
$PswdValLabel.Location = '400,300'
$PswdValLabel.AutoSize = $true

$SidMenu.Controls.Add($PswdValLabel)

$RandomPswdButton = New-Object System.Windows.Forms.Button
$RandomPswdButton.Location = '570,300'
$RandomPswdButton.Text = 'Random'
$RandomPswdButton.AutoSize = $true

$SidMenu.Controls.Add($RandomPswdButton)

$okbutton = New-Object System.Windows.Forms.Button
$cancelbutton = New-Object System.Windows.Forms.Button
$validatebutton = New-Object System.Windows.Forms.Button
$resetbutton = New-Object System.Windows.Forms.Button

$okbutton.AutoSize = $true
$cancelbutton.AutoSize = $true
$validatebutton.AutoSize = $true
$resetbutton.AutoSize = $true

$resetbutton.Text = 'Reset'
$resetbutton.Location = '260,350'

$validatebutton.Text = 'Validate'
$validatebutton.Location = '365,350'

$okbutton.Text = 'ok'
$okbutton.Location = '470,350'

$cancelbutton.Text = 'cancel'
$cancelbutton.Location = '575,350'

$SidMenu.Controls.Add($resetbutton)
$SidMenu.Controls.Add($validatebutton)
$SidMenu.Controls.Add($okbutton)
$SidMenu.Controls.Add($cancelbutton)

$NullLabel = New-Object System.Windows.Forms.Label
$NullLabel.Location = '50,300'
$NullLabel.ForeColor = 'Red'

$SidMenu.Controls.Add($NullLabel)

$SendMailButton = New-Object System.Windows.Forms.Button
$SendMailButton.Text = "Send Email"
$SendMailButton.Location = '130,350'
$SendMailButton.AutoSize = $true

$SidMenu.Controls.Add($SendMailButton)

#Logic Part

#Reset Button Logic
$resetbutton.Add_Click({
    $netidtxtbox.Clear()
    $Desctxtbox.Clear()
    $Tsktxtbox.Clear()
    $Pswdtxtbox.Clear()
    $Wstntxtbox.Clear()
    $Ownertxtbox.Clear()
    $PswdValLabel.Text = ""
    $validateuserlabel.Text = ""
    $ValidateOwner.Text = ""
    })

#Password Logic

$validatebutton.add_click({
    #validate user
    $User = $(try {Get-ADUser $netidtxtbox.Text -ErrorAction Stop} catch {$null})
    if($User -ne $null) {
    $validateuserlabel.Text = "User already exists in AD"
    $validateuserlabel.ForeColor = 'Red'
    }else{
    $validateuserlabel.Text ="User is available for creation"
    $validateuserlabel.ForeColor = 'green'}
    
    
    #validate Password
    
    #Setting minimum password length to 8 characters and adding password complexity.
        $Password = ConvertTo-SecureString $Pswdtxtbox.Text -AsPlainText -Force
        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)
        $Complexity = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
                
        if(($Complexity -ne $null) -and ($Complexity -cmatch '[a-z]') -and ($Complexity -cmatch '[A-Z]') -and ($Complexity -match '\d') -and ($Complexity.length -ge 8) -and ($Complexity -match '!|@|#|%|^|&|$'))
        {
            $PswdValLabel.Text = "Strong password"
            $PswdValLabel.ForeColor = 'Green'
        }
            else
        {
            $PswdValLabel.Text = "Invalid password"
            $PswdValLabel.ForeColor = 'Red'
        }
        
    #Validate Owner
    $owner = $(try {Get-ADUser $Ownertxtbox.Text -ErrorAction Stop} catch {$null})
    if($owner -ne $null){
    $ValidateOwner.Text = "Owner Net id is valid"
    $ValidateOwner.ForeColor = "Green"
    }else{
    $ValidateOwner.Text = "Owner Net id is not valid"
    $ValidateOwner.ForeColor = 'Red'
    }
    })
#Random Password generator
Function GenerateStrongPassword ([Parameter(Mandatory=$true)][int]$PasswordLenght)
{
Add-Type -AssemblyName System.Web
$PassComplexCheck = $false
do {
$newPassword=[System.Web.Security.Membership]::GeneratePassword($PasswordLenght,1)
If ( ($newPassword -cmatch "[A-Z\p{Lu}\s]") `
-and ($newPassword -cmatch "[a-z\p{Ll}\s]") `
-and ($newPassword -match "[\d]") `
-and ($newPassword -match "[^\w]")
)
{
$PassComplexCheck=$True
}
} While ($PassComplexCheck -eq $false)
return $newPassword
}

$RandomPswdButton.Add_Click({

$Pswdtxtbox.Text = GenerateStrongPassword(8) })
    
#ID CREATION LOGIC
#fixed variable
$OU = "OU=SID ServiceID,OU=Service Accounts,OU=Administration,OU=APTIV,DC=aptiv,DC=com"
$domain = "aptiv.com"

$SidMenu.CancelButton = $cancelbutton

$okbutton.add_click({

#Textbox variable assignment
$Sam = $netidtxtbox.Text
$Description = $Desctxtbox.Text
$Notes = $Tsktxtbox.Text
$WorkStations = $Wstntxtbox.Text
$Pass = $Pswdtxtbox.Text


if(($netidtxtbox.Text -ne "") -and ($Desctxtbox.Text -ne "") -and ($Tsktxtbox.Text -ne "") -and ($Pswdtxtbox.Text -ne ""))
    
    {
        #if account does not exists then create a new one
        try{
        $Password = ConvertTo-SecureString "$Pass" -AsPlainText -Force
         New-ADUser -SamAccountName $Sam `
            -UserPrincipalName "$Sam@aptiv.com" `
            -Name $Sam `
            -GivenName $Sam `
            -Surname "" `
            -Enabled $True `
            -ChangePasswordAtLogon $False `
            -CannotChangePassword $True `
            -PasswordNeverExpires $True `
            -DisplayName $Sam `
            -Description $Description `
            -Path $OU `
            -AccountPassword (convertto-securestring $Password -AsPlainText -Force)

            $pas = $Pswdtxtbox.Text
            Write-Information "Created $sam with below information :"
            Write-Information "ou = $ou" 
            Write-Information "Description = $Description" 
            Write-Information "Password = $pas"

            }catch
            {
            $validateuserlabel.Text = "User already exists in AD"
            exit
            }
            
          Sleep 3

          # Add SID to groups
            Add-ADGroupMember -Identity "serviceIDS" -Members $Sam
            Add-ADGroupMember -Identity "AUDIT_SPECIAL ID" -Members $Sam
            Add-ADGroupMember -Identity "Avecto_StandardUser" -Members $Sam
            Write-Information "Added $Sam to serviceIDS, AUDIT_SPECIAL ID & Avecto_StandardUser Groups "
          Sleep 1

          # adding Notes Tab
          Set-ADUser -Identity $Sam -ErrorAction Stop -Replace @{'info'= $Notes}
          Write-Information "Task Number = $Notes"
          
          Sleep 1

          #adding workstations
          try{
          foreach($workstation in $WorkStations){Set-ADUser -Identity $Sam -LogonWorkstations $workstation}
          Write-Information "Workstations added = $WorkStations"
          }catch{Write-Information "No Workstation added"}
          
          Sleep 1

          #Updating Managedby
          Set-ADUser -Identity $Sam -ErrorAction Stop -Manager $Ownertxtbox.Text
          $own = $Ownertxtbox.Text
          Write-Information "Owner = $Own"

          Sleep 1 
          [void][System.Windows.MessageBox]::Show( "All changes have been implemented successfully ", "Script completed", "OK", "Information" )
          }
          else
          { 
          [void][System.Windows.MessageBox]::Show( "Please fill all the details ", "Error", "OK", "Information" )
          }
})

#Send Eamil Button Function
Function Send_Email{
$smtpServer = "smtp.aptiv.com"
$smtpFrom = "Special_ID_No_Reply@aptiv.com"
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

$to = (Get-ADUser $Ownertxtbox.Text -Properties *|select UserPrincipalName).UserPrincipalName
$bcc = "hcl_Aptiv_AD@hcl.com"
$Subject = $Tsktxtbox.Text

$body = 'Classification: Confidential' + "`n" + "`n"
$body += 'Hi ' + $to.split(".",2)[0] + ',' + "`n"+ "`n"
$body += 'Please use "' + $Pswdtxtbox.Text + '" as password for ' + $netidtxtbox.Text + '.' + "`n"+ "`n"
$body += 'Regards,' + "`n"
$body += 'AD Team Automation'

Send_Email

[void][System.Windows.MessageBox]::Show( "Mail Sent Successfully ", "Mail", "OK", "Information" )
Write-Information "shared password information with $to bcc $bcc subject = $Subject"
})

$SidMenu.ShowDialog()
Stop-Transcript

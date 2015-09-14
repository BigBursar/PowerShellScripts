Import-Module ActiveDirectory

# CSV file containing user details
$ImportedUsers = Import-Csv "inputs\HR-Users-Export.csv"

# Arrays to hold found and missing users
$FoundUsers = @()
$MissingUsers = @()
$NoEmailUsers = @()

[string]$EmailBody = ""

#Loop through each user in the spreadsheet
ForEach ($ImportedUser in $ImportedUsers)
{ 
    $GivenName = $ImportedUser.'Known As'
    $Surname = $ImportedUser.'Last Name'
    $Email=$ImportedUser.Email
    $Manager=$ImportedUser.Manager
    $DeskNum = $ImportedUser.'Tel. Extension'
    $Location=$ImportedUser.Location
    $Department=$ImportedUser.Department

    # If the spreadsheet entry has an email address
    if (![string]::IsNullOrEmpty($Email)) {

        #Find matching AD account
        if (Get-ADUser -Filter { Mail -eq $Email  }) {
            # Get AD account fields
            $ADUser = Get-ADUser -Filter { Mail -eq $Email  }

            # Check for empty fields and fill with default values
            if ([string]::IsNullOrEmpty($Department)) { $Department = " " }

            if ([string]::IsNullOrEmpty($Location)) { $Location = " " }

            if ([string]::IsNullOrEmpty($DeskNum)) { 
                $FormattedDeskNum = "n/a"
            } else {
                # Format Desk Phone number
                $DeskNum=$ImportedUser.'Tel. Extension'.Substring($ImportedUser.'Tel. Extension'.Length-3,3)
                $FormattedDeskNum = "01954 234 $DeskNum"                
            }
    
            if ([string]::IsNullOrEmpty($ImportedUser.'Work Telephone')) {
                $FormattedMobileNum = "n/a"
            } else {
                # Format Mobile Phone Number
                $FormattedMobileNum = "{0:+############}" -f $ImportedUser.'Work Telephone'
            }

            # Update fields
            # Job Titles in the spreadsheet are currently prefixed with garbage data so they are being ignored
            # Once the job titles in the HR system are correct, uncomment the line below, and comment the line below that
            # Set-ADUser $ADUser -Description $ImportedUser.'Job Title' -Title $ImportedUser.'Job Title' -Department $Department -Company $Location `
            Set-ADUser $ADUser -Department $Department -Company $Location `
                -OfficePhone $FormattedDeskNum -MobilePhone $FormattedMobileNum

            # Get and update Manager
            if ($ADManager=Get-ADUser -Filter { DisplayName -eq $Manager  }) {
                Set-ADUser $ADUser -Manager $ADManager.SamAccountName
            } else {
            # Manager account doesn't exist, so do nothing
            }

            # Update list of found users
            $FoundUsers += ,@($GivenName, $Surname, $Email) 
        } else {
            # HR User doesn't exist in AD, so update missing users
            $MissingUsers += ,@($GivenName, $Surname, $Email) 
        }
    } else {
        # HR User has no email address, so update no email users
        $NoEmailUsers += ,@($GivenName, $Surname, $Email)
    }
}

# Create Report Email
$EmailBody = "<p><font color = 'green'>The following users were found and updated</font></p>"
ForEach ($User in ($FoundUsers | Sort-Object)){
#    Write-Host "$User found and updated" -ForegroundColor "Green"
    $EmailBody += "$User<br/>"
}
$EmailBody += "<hr/><p><font color = 'red'>The following users are in the CSV file, but NOT in AD</font><br/></p>"
ForEach ($User in ($MissingUsers | Sort-Object)){
#    Write-Host "$User missing from Active Directory!`r`n" -ForegroundColor "Red"
    $EmailBody += "$User<br />"
}
$EmailBody += "`<hr/><p><font color = 'blue'>The following users have no email address in the CVS file</font></p>"
ForEach ($User in ($NoEmailUsers | Sort-Object)){
#    Write-Host "$User has no email address in HR!`r`n" -ForegroundColor "Yellow"
    $EmailBody += "$User <br />"
}

# Send report email - update addresses to valid domains!!
send-mailmessage -to "admin@contoso.com" -from "AD Update Script <adupdatescript@contoso.com>" -subject "AD Update Script Results" -smtpServer mail.contoso.com `
    -body $EmailBody -BodyAsHtml
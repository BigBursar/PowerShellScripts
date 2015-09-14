# Get login credentials to an Exchange server
$cred = Get-Credential -Message "Log in with valid Username and Password"
# Update the -ConnectionURI field with the name of an Exchange CAS server
$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://<servername>/powershell -Credential $cred
Import-PSSession $session

# Import all active AD Users
## Update the variable below with the domain to look for e.g. '*contoso*'
$ADUsers = Get-MailContact -resultsize unlimited -Filter { ExternalEmailAddress -like '*contoso*' } -SortBy Name

# Create HTML Table
$HTML = "<table><tr><td><strong>Name</strong></td><td><strong>E-Mail</strong></td><td><strong>Telephone</strong></td><td><strong>Mobile</strong></td>"
$HTML+= "<td><strong>Skype</strong></td><td><strong>Title</strong></td><td><strong>Office</strong></td></tr>"

#Loop through each user
ForEach ($User in $ADUsers) {
    # Get the required bits of information
    $email = $User.ExternalEmailAddress.Trim("SMTP:")
    $Contact = Get-Contact -Identity $email
    $name = $User.Name
    $title=$Contact.Title
    $phone=$Contact.Phone
    $mobile=$Contact.MobilePhone
    $skype=$User.CustomAttribute1
    $office=$Contact.Office

    # Output in HTML format
    $HTML += "<tr>"
    $HTML += "<td>$name</td>"
    $HTML += "<td><a href=""mailto:$email"">$email</a></td>"
    $HTML += "<td>$phone</td>"
    $HTML += "<td>$mobile</td>"
    $HTML += "<td>$skype</td>"
    $HTML += "<td>$title</td>"
    $HTML += "<td>$office</td>"
	$HTML += "</tr>"
}

$HTML += "</table>"

# Output the results into a file
$HTML | Out-File "outputs/contacts.html"
# Start of Settings

$GLOBAL:PathToXAML = "C:\Dropbox\github_repos\Archiver"
$GLOBAL:user = $env:USERNAME
$GLOBAL:date = get-date

$GLOBAL:onSiteStorage = "C:\gatewayTest\onSiteStorage\"
$GLOBAL:offSiteStorage = "C:\gatewayTest\offSiteStorage\"
$GLOBAL:latestLogs = "C:\gatewayTest\LatestArchive\"

# Please Specify the SMTP server address
$GLOBAL:smtpServer = "relay.nhs.uk"
# Please specify the email address who will send the report
$GLOBAL:emailFrom ="billy.westbury@nhs.net"
# Please specify the email address who will receive the report
$GLOBAL:emailToList = "billy.westbury@nhs.net"
# Please specify an email subject
$GLOBAL:emailSubject="Report: Automated Archive Creation Success: $date"
# Please specify an email body
$GLOBAL:emailHeader = @"
	
Audit Files Archive Report for $Date
	
The following Archive has been succesfully copied:

"@        
$GLOBAL:emailFooter = @"
	
The files have been succesfully backed up to both onsite and offsite storage servers.

This was carried out by $user
	
		
"@

# Use the following item to define if an email report should be sent once completed
$GLOBAL:sendEmail = $true
# If you would prefer the Excel file as an attachment then enable the following:
$GLOBAL:SendAttachment = $false
# The path to the excel report
$GLOBAL:excelFile = "C:\gatewayTest\ArchiveReport\ArchiveDetails.xlsx"

# End of Settings


# Start of Settings

$GLOBAL:user = $env:USERNAME
$GLOBAL:date = get-date

$GLOBAL:onSiteStorage = "C:\gatewayTest\onSiteStorage\"
$GLOBAL:offSiteStorage = "C:\gatewayTest\offSiteStorage\"
$GLOBAL:latestLogs = "C:\gatewayTest\LatestArchive\"

# Please Specify the SMTP server address
$GLOBAL:smtpServer = ""
# Please specify the email address who will send the report
$GLOBAL:emailFrom =""
# Please specify the email address who will receive the report
$GLOBAL:emailToList = ""
# Please specify an email subject
$GLOBAL:emailSubject="Report: Automated Archive Creation Success: $date"
# Please specify an email body
$GLOBAL:emailBody = @"
	
Audit Files Archive Report for $Date
	
The following files have been succesfully copied:

		$GLOBAL:newFilelist
	
The files have been succesfully backed up to both onsite and offsite storage servers.

This was carried out by $user
	
		
"@

# Use the following item to define if an email report should be sent once completed
$GLOBAL:sendEmail = $true
Write-Debug $GLOBAL:sendEmail
# If you would prefer the Excel file as an attachment then enable the following:
$GLOBAL:SendAttachment = $false
# The path to the excel report
$GLOBAL:excelFile = "C:\gatewayTest\ArchiveReport\ArchiveDetails.xlsx"

# End of Settings


# This script analyses a user's mailbox and if it is below a certain threshold it will convert it to a shared mailbox and unlicense the account
$Session = Get-PSSession
$Emails = Get-Content "Path to csv"
$Mailto = "Person's Name<Person@Company.com"
$Mailfrom = "Person's Name<Person@Company.com"
$SubjectMessage = "Subject of e-mail"
$BodyMessage = "Body of e-mail"
$SMTPServer = "smtpserver.company.com"
$O365License = "Type of office 365 License"
If (-Not($Session.ComputerName -eq "outlook.office365.com" -and $Session.State -eq "Opened")) # Tests for already existing powershell session to Microsoft on-line (MSOL)
	{
	$UserCredential = Get-Credential
	Connect-MSOLservice -Credential $UserCredential
	$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
	Import-PSSession $Session -AllowClobber
	}
ForEach ($Email in $Emails) #Loops through CSV and determines size of each mailbox
	{
	$MailboxSizeTemp=(Get-MailboxStatistics -Identity $Email).TotalItemSize
	$MailboxSize=($MailboxSizeTemp -Replace '(\,|^[^\(]*\(| bytes\))','')
	[int64]$intMailboxSize = [convert]::ToInt64($MailboxSize, 10)
	$MaxSize = 10737418240
	If ((Get-Mailbox $Email).IsShared -eq $False) #Checks to make sure that the mailbox isn't already shared
		{
		If ($intMailboxsize -lt $MaxSize) #If the mailbox is less than 10GB it will convert it to a shared mailbox and unlicense the account
			{	
			Set-Mailbox $Email -Type Shared -ProhibitSendReceiveQuota 10GB -ProhibitSendQuota 9.75GB -IssueWarningQuota 9.5GB
			Set-MSOLUserLicense -UserPrincipalName $Email -RemoveLicenses $O365License
			}
		Else #If the mailbox is too large, it will send an e-mail to IT customer service announcing the the mailbox is too large to convert.
			{
			$intMailboxSize = [math]::round($intMailBoxSize/1Gb,2)
			Send-MailMessage -To $Mailto -From $MailFrom -Subject $SubjectMessage -Body $BodyMessage -SMTPserver $SMTPServer
			}
		}
	}
Get-PSSession | Remove-PSSession1
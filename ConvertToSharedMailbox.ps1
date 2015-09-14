# This script analyses a user's mailbox and if it is below a certain threshold it will convert it to a shared mailbox and unlicense the account
$Session = Get-PSSession
$Emails = Get-Content C:\Users\borda01\Test1.csv
If (-Not($Session.ComputerName -eq "outlook.office365.com" -and $Session.State -eq "Opened")) # Tests for already existing powershell session to Microsoft on-line (MSOL)
	{
	$UserCredential = Get-Credential
	connect-msolservice -credential $UserCredential
	$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
	Import-PSSession $Session -AllowClobber
	}
ForEach ($Email in $Emails) #Loops through CSV and determines size of each mailbox
	{
	$MailboxSizeTemp=(Get-MailboxStatistics -identity $Email).totalItemSize
	$MailboxSize=($MailboxSizeTemp -Replace '(\,|^[^\(]*\(| bytes\))','')
	[int64]$intMailboxSize = [convert]::ToInt64($MailboxSize, 10)
	$Maxsize = 10737418240
	If ($intMailboxsize -lt $maxsize) #If the mailbox is less than 10GB it will convert it to a shared mailbox and unlicense the account
		{	
		Set-Mailbox $Email -Type Shared -ProhibitSendReceiveQuota 10GB -ProhibitSendQuota 9.75GB -IssueWarningQuota 9.5GB
		Set-MSOLUserLicense -UserPrincipalName $Email -RemoveLicenses pvsw:ENTERPRISEWITHSCAL
		}		
	Else #If the mailbox is too large, it will send an e-mail to IT customer service announcing the the mailbox is too large to convert.
		{
		$intMailboxSize = [math]::round($intMailBoxSize/1Gb,2)
		Send-Mailmessage -to "Desktop Support <IT.Customer.Service@actian.com>" -from "Term Script <IT.Customer.Service@actian.com>" -Subject "Mailbox too large to convert to shared mailbox" -Body "$Email is unabled to be archived because it is $intMailboxSize GB.  The largest convertable is 10GB." -SMTPserver smtp.actian.com
		}
	}
get-pssession | remove-pssession
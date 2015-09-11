$Sessions= Get-PSSession
ForEach ($Session in $Sessions)
	{
	If ($Session.ComputerName -eq "outlook.office365.com" -and $Session.state -eq "Opened")
		{
		#Write-Host "Connected to MSOL service" $session.ComputerName $session
		$Emails = Get-Content C:\Users\borda01\termed_mails2.csv
		$MailboxSizeTemp=(Get-MailboxStatistics -identity dan.borges@actian.com).totalItemSize
		$MailboxSize=($MailboxsizeTemp -Replace '(\,|^[^\(]*\(| bytes\))','')
		[int]$intMailboxsize = [convert]::ToInt32($MailboxSize, 10)
		$Maxsize = 10737418240
		If ($intMailboxsize -lt $maxsize)
			{
			#Write-Host $intMailboxsize" is smaller than " $Maxsize
			ForEach ($Email in $Emails)
				{
				Set-Mailbox $Email -type Shared -ProhibitSendReceiveQuota 10GB -ProhibitSendQuota 9.75GB -IssueWarningQuota 9.5GB
				Set-MSOLUserLicense -UserPrincipalName $Email -RemoveLicenses pvsw:ENTERPRISEWITHSCAL
				}
			}
		Else
			{
			#Write-Host $intMailboxsize "is not smaller than " $Maxsize
			Send-Mailmessage -to "Benita Trevino <benita.trevino@actian.com>" -from "Dan Borges <dan.borges@actian.com>" -Subject "Mailbox too large to convert to shared mailbox" -Body "$Email is unabled to be archived because it is $intMailboxSize" -SMTPserver smtp.actian.com
			}		
		}
	Else
		{
		#Write-Host "Do not connected to MSOL service" $Session.Computername
		$UserCredential = Get-Credential
		connect-msolservice -credential $UserCredential
		$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
		Import-PSSession $Session -AllowClobber
		$Emails = Get-Content C:\Users\borda01\termed_mails2.csv
		$MailboxSizeTemp=(Get-MailboxStatistics -identity dan.borges@actian.com).totalItemSize
		$MailboxSize=($MailboxsizeTemp -Replace '(\,|^[^\(]*\(| bytes\))','')
		[int]$intMailboxsize = [convert]::ToInt32($MailboxSize, 10)
		$Maxsize = 10737418240
		If ($intMailboxsize -lt $maxsize)
			{
			#Write-Host $intMailboxsize" is smaller than " $Maxsize
			ForEach ($Email in $Emails)
				{
				Set-Mailbox $Email -type Shared -ProhibitSendReceiveQuota 10GB -ProhibitSendQuota 9.75GB -IssueWarningQuota 9.5GB
				Set-MSOLUserLicense -UserPrincipalName $Email -RemoveLicenses pvsw:ENTERPRISEWITHSCAL
				}
			}
		Else
			{
			#Write-Host $intMailboxsize "is not smaller than " $Maxsize
			Send-Mailmessage -to "Benita Trevino <benita.trevino@actian.com>" -from "Dan Borges <dan.borges@actian.com>" -Subject "Mailbox too large to convert to shared mailbox" -Body "$Email is unabled to be archived because it is $intMailboxSize" -SMTPserver smtp.actian.com
			}
		}
	}
	
#Regular expression

#Remove commas from string:
#$String -Replace '\,', ''  
#\ = escape
# single quote  escape comma single quote comma space single quote single quote
#$String -Replace '^[^\(]*\(',''
#$String -Replace ' bytes\)', ''
#$String -Replace  '\,', '' -Replace '^[^\(]*\(','' -Replace ' bytes\)', ''

#'(\,|^[^\(]*\(| bytes\))',''
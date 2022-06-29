# Heath Stewart - 31 March 2020
# Updated 20 September 2021
# This script will remove phishing emails from user inboxes as long as you know about them.
# It is done by going through all exchange mailboxes.
# The accepted inputs are the subject (all or partial) and the sender domain - all are case-insensitive
#
# 20201110 Update: The results of the $delete option are tee'd out to a file so that potential
# victims of a phishing campaign can be monitored. Results are emailed at the end.
#
# 20210920 Update: Added a user count to the body of the email to measure impacted mailboxes quickly.


# set acceptable parameters for script
param(
    [string] $subject,
    [string] $from,
    [switch] $delete
)

# Need to create and import a new powershell session to be able to connect to and run commands from exchange locally.
# ****NOTE: Uncomment the following 2 lines if you are running this via PS remoting on your PC rather than exchange server
#$exsession = New-PSSession -ConfigurationName microsoft.exchange -ConnectionUri http://<YOUR EXCHANGE SERVER>/powershell
#Import-PSSession $exsession

# setup the email variables for the report
$today = Get-Date  
$reportSubject = "Phishing Recipients Report - " + $today 
$priority = "Normal" 
$smtpServer = "<YOUR EXCHANGE SERVER>" 
$emailFrom = "<FROM ADDRESS>" 
$emailTo = "<TO ADDRESS>"
$body = "Check the attached accounts for suspicious activity regarding the following suspicious email.`n
    SUBJECT: $subject
    FROM DOMAIN: $from`n
The messages have been deleted with the command below, but check for compromise of affected user accounts!`n
    c:\scripts\powershell\shellphish.ps1 -subject `"$subject`" -from `"$from`" -delete`n`n"

# initialize files to store the results
$results = "c:\scripts\powershell\results.txt"

# get all the mailboxes
$mailboxes = Get-Mailbox -ResultSize unlimited

# verify that the inbox contains the message in question, default is log it in your mailbox
# otherwise, delete the message
if($delete) {
    write-output "Affected Mailboxes:" > $results
    write-host "Preparing to delete messages with subject `"$subject`""
    foreach($box in $mailboxes) {
        Search-Mailbox -identity $box.DisplayName -SearchQuery "subject:`"$subject`" AND from:$from" -DeleteContent -Force | 
            Where-Object -Property ResultItemsCount -NE -value "0" | Tee-Object -filepath $results -Append
    }
    
    # count the number of affected mailboxes and store it for email body
    $count = select-string -Path $results -pattern "Identity" | measure | select-object count
    $numusers = $count.Count

    # return a list and send email of affected mailboxes
    write-host "Affected mailboxes in `"$results`""
    Send-MailMessage -To $emailTo -Subject $reportSubject -Attachments $results -Body "$body Number of users affected: $numusers" -SmtpServer $smtpServer -From $emailFrom -Priority $priority
}

# if -delete is not used, the mailboxes are only checked for the phishing email
# and they are reported to your mailbox in a Search4Phish folder
else {
    foreach($box in $mailboxes) {
        Write-Host "Searching $box for `"$subject`""
        Search-Mailbox -identity $box.DisplayName -SearchQuery "subject:`"$subject`" AND from:$from" -TargetMailbox "<YOUR EMAIL ADDRESS>" -TargetFolder "Search4Phish" -LogOnly
    }
    write-host "Check Inbox/Search4Phish for messages matching the search criteria."
}

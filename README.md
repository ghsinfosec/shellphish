# shellphish
Email removal script for on-prem Exchange I wrote a few years ago. Most orgs are going 365 now, so I don't think it's really necessary. But if you're using on-prem Exchange it's useful.

Depending on the number of mailboxes it can take a little while to run, but it works.

Also I'm no powershell guru, but sometimes quick and dirty is all you need.

# usage
Run it from the directory you installed it like any Windows script. Examples below.

## Example - Query user mailboxes only
Specify the email subject and sender address to find and report end-user mailboxes that received an email meeting the `-subject` and `-from` criteria. Only users that have emails matching those criteria are reported.

`c:\scripts\powershell\shellphish.ps1 -subject "Verify your mailbox or be shutdown!!" -from "somehacker@gmail.com"`

## Example - DELETE all messages specified by the arguments
Same as the above example, but with the `-delete` option which will delete the messages in all affected mailboxes and email a results file of those affected mailboxes. BE CAREFUL USING THIS OPTION!

`c:\scripts\powershell\shellphish.ps1 -subject "Verify your mailbox or be shutdown!!" -from "somehacker@gmail.com" -delete`

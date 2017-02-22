A powershell script to create emails in outlook

Relies on a set of comma separated files tokens to populate and send emails. Emails will not be sent but instead saved as drafts in Outlook for review.

Users:

1. Right click "SendEmail.ps1" and select "Run With PowerShell"
2. Follow instructions in the script, select whatever applies to you.
3. Open draft-folder in Outlook, review the emails and send them.


For admins:

1. Make sure the information in the "signature_*.csv" and "vouchers_*.csv" is correct, the headers is treated as tokens
2. Edit the OFT-files correctly inserting tokens in the following form "$TokenName$" (tokens are in the first line of the csv-files
3. Tell Outlook to "ignore all" on any spelling errors regarding tokens.
4. Scroll down and decide whether you want to have a signature present, if not, remove it.
5. Save the files in a oft-format.
6. Distribute in a common folder on Box (just a suggestion)

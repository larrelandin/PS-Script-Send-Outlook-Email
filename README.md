A powershell script to create emails in outlook

Uses a csv-file with tokens to populate and send emails. Emails will not be sent but instead saved as drafts in Outlook for review.

1. Make sure the information in "Recipients.txt" and "GlobalTokenReplacements.txt" is correct, the headers is treated as tokens
2. Edit the OTF-files correctly inserting tokens in the following form "[TokenName]"
3. Right click "SendEmail.ps1" and select "Run With PowerShell"
4. Follow instructions in the script.

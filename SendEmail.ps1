### Script source available here: https://github.com/larrelandin/PS-Script-Send-Outlook-Email ###

### Filter for voucher-files e.g. vouchers_MilanFeb2017.csv ###
$voucherFilter = "vouchers_*.csv"

### Filter for signature-files e.g. signature_lal.csv
$signatureFilter = "signature_*.csv"


### Tokens and separators ###
$StartTokenIdentifier = '$'
$EndTokenIdentifier = '$'
$CSVDelimiter =';'


####################
### Script start ###
####################

### General path to the script and files ###
$mainPath = (Split-Path -Path $MyInvocation.MyCommand.Definition -Parent) + '\'

Clear-Host

###

$signatures = Get-ChildItem -Path $mainPath -Filter $signatureFilter
$signaturePath = $signatures | Out-GridView -Title "Select the appropriate signature file" -PassThru

###

$vouchers = Get-ChildItem -Path $mainPath -Filter $voucherFilter
$voucherPath = $vouchers | Out-GridView -Title "Select the appropriate voucher file" -PassThru

###

$mailTemplates = Get-ChildItem -Path $mainPath -Filter "*.oft"
$mailTemplate = $mailTemplates | Out-GridView -Title "Select an email template" -PassThru

write-host "You selected the following Email Template:"
write-host $mailTemplate -ForegroundColor Green
write-host "`n"

###

### Importing Signature Tokens ###
$SignatureFile = Import-Csv "$mainPath$SignaturePath" -Delimiter $CSVDelimiter
write-host "These are your local tokens:"
write-host $signaturePath -ForegroundColor Green
Write-Host ($SignatureFile | Format-Table | Out-String) -ForegroundColor Green

### Importing the CSV ###
$recipients = Import-Csv "$mainPath$voucherPath" -Delimiter $CSVDelimiter

Write-Host "This is your list with vouchers:"
write-host $voucherPath -ForegroundColor Green
Write-Host ($recipients | Format-Table | Out-String) -ForegroundColor Green

[ValidateSet('y','n')]$Answ1 = Read-Host "Is the mapped information correct? [y]es or [n]o?"

if($Answ1 -eq 'y')
{

    ### Output a list of emails for verification ###
    foreach($rec in $recipients)
    {
        Write-Host $rec.Email -ForegroundColor Green
    }
    Write-Host " "
    [ValidateSet('y','n')]$Answ2 = Read-Host "Emails seem correct? [y]es or [n]o?"
    if($Answ2 -eq 'y')
    {

        if($null -ne (Get-Process -name outlook -ErrorAction SilentlyContinue))
        {
            Write-Host -ForegroundColor Red "Outlook seems to be running, you need to shut it off before continuing"
            Write-Host -ForegroundColor Red "Press any key to continue..."
            $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
        }
        
        foreach($rec in $recipients)
        {
            ### Send Email using Outlook Client ###
            $Outlook = New-Object -ComObject Outlook.Application
            $Mail = $Outlook.CreateItemFromTemplate("$mainPath$mailTemplate")

            ### Replace global tokens in html-body ###
            foreach($gt in $SignatureFile.GetEnumerator())
            {
                $Mail.HTMLBody = $Mail.HTMLBody.Replace("$StartTokenIdentifier$($gt.Name)$EndTokenIdentifier",$($gt.Value))
            }

            ### Replace individual tokens in html-body ###
            foreach($it in ($rec | Get-Member -MemberType NoteProperty))
            {
                $token = "$StartTokenIdentifier$($it.Name)$EndTokenIdentifier"
                #Write-Host $token
                $value = $rec.$($it.Name)
                #Write-Host $value
                $Mail.HTMLBody = $Mail.HTMLBody.Replace($token,$value)
            }


            $Mail.To = $rec.Email
            $Mail.Save()
            ##$Mail.Send()

            Write-Host "DONE: " $rec.Email -ForegroundColor Cyan
        }
    }
}

### Trainer Name ###
$TrainerName = "Larre Ländin"

### General path to the script and files ###
$mainPath = "C:\Users\lal\Box Sync\Training\SXSD\SXSD Email sendout\"

### CSV with vouchers ###
$voucherPath = "csv-file.txt"

### Exam Email ###
$HTMLExamEmailPath = "Developer exam access.htm"

### Evaluation Email ###
$HTMLEvalEmailPath = "Developer eval access.htm"

### Email subject lines ###
$ExamEmailSubject = "SXSD Developer Exam Access"
$EvalEmailSubject = "SXSD Developer Evaluation Access"

### These tokens will be replaced by the corresponding value ###
$CommonTokenReplacements = @{"[TrainerName]" = $TrainerName}

### These tokens will be replaced by the value from the CSV ###
$IndividualTokenReplacements = @{"[ReceiverFirstName]" = "First Name"; "[ReceiverLastName]" = "Last Name"; "[VoucherExamCode]" = "Voucher Test"; "[VoucherEvalCode]" = "Voucher Eval"; "[PartnerName]" = "Partner"}


####################
### Script start ###
####################
Clear-Host

[ValidateSet('x','v')]$AnswExam = Read-Host "Do you want to generate e[x]am email or e[v]aluate email?"
if($AnswExam -eq 'x')
{
    $HTMLSelectedEmailPath = $HTMLExamEmailPath
    $EmailSubject = $ExamEmailSubject
}
elseif($AnswExam -eq 'v')
{
    $HTMLSelectedEmailPath = $HTMLEvalEmailPath
    $EmailSubject = $EvalEmailSubject
}
Clear-Host

$CommonTokenReplacements
$IndividualTokenReplacements

### Importing the CSV ###
$participants = Import-Csv "$mainPath$voucherPath"

Write-Host ($participants | Format-Table | Out-String)

[ValidateSet('y','n')]$Answ1 = Read-Host "Is the mapped information correct? [y]es or [n]o?"

if($Answ1 -eq 'y')
{
    Clear-Host

    ### Output a list of emails for verification ###
    foreach($part in $participants)
    {
        Write-Host $part.Email
        
    }
    Write-Host " "
    [ValidateSet('y','n')]$Answ2 = Read-Host "Emails seem correct? [y]es or [n]o?"
    if($Answ2 -eq 'y')
    {
        Clear-Host
        ### Replace common tokens in html-email ###
        $HTMLMaster = Get-Content "$mainPath$HTMLSelectedEmailPath"

        foreach($h in $CommonTokenReplacements.GetEnumerator())
        {
            #Write-Host "$($h.Name): $($h.Value)"
            $HTMLMaster = $HTMLMaster.Replace($($h.Name),$($h.Value))
        }

        foreach($part in $participants)
        {
            ### Replace individual tokens in html-email ###

            #Create a copy for every email
            $HTMLMessage = $HTMLMaster

            foreach($h in $IndividualTokenReplacements.GetEnumerator())
            {
                #Write-Host "$($h.Name): $($h.Value)"
                $HTMLMessage = $HTMLMessage.Replace($($h.Name),$part.$($h.Value))
            }
            
            ### Send Email using Outlook Client ###
            $Outlook = New-Object -ComObject Outlook.Application
            $Mail = $Outlook.CreateItem(0)

            $Mail.To = $part.Email
            $Mail.Subject = $EmailSubject
            $Mail.HTMLBody = [string]$HTMLMessage
            $Mail.Save()
            #$Mail.Send()

            Write-Host "DONE: " $part.Email
        }
    }
}

### Script source available here: https://github.com/larrelandin/PS-Script-Send-Outlook-Email ###

### General path to the script and files ###
$mainPath = (Split-Path -Path $MyInvocation.MyCommand.Definition -Parent) + '\'

### CSV with vouchers ###
$voucherPath = "Recipients.txt"

### Exam Email ###
$HTMLExamEmailPath = "Developer exam access.oft"

### Evaluation Email ###
$HTMLEvalEmailPath = "Developer evaluation access.oft"

### Tokens ###
$StartTokenIdentifier = '['
$EndTokenIdentifier = ']'

### File with Global Token Replacements for all emails ###
$GlobalTokenReplacementsFile = 'GlobalTokenReplacements.txt'

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

### Importing Global Tokens ###
$GlobalTokenReplacements = Import-Csv "$mainPath$GlobalTokenReplacementsFile"
Write-Host ($GlobalTokenReplacements | Format-Table | Out-String)

### Importing the CSV ###
$recipients = Import-Csv "$mainPath$voucherPath"

Write-Host ($recipients | Format-Table | Out-String)

[ValidateSet('y','n')]$Answ1 = Read-Host "Is the mapped information correct? [y]es or [n]o?"

if($Answ1 -eq 'y')
{
    Clear-Host

    ### Output a list of emails for verification ###
    foreach($rec in $recipients)
    {
        Write-Host $rec.Email
        
    }
    Write-Host " "
    [ValidateSet('y','n')]$Answ2 = Read-Host "Emails seem correct? [y]es or [n]o?"
    if($Answ2 -eq 'y')
    {
        Clear-Host
        
        foreach($rec in $recipients)
        {
            ### Send Email using Outlook Client ###
            $Outlook = New-Object -ComObject Outlook.Application
            $Mail = $Outlook.CreateItemFromTemplate("$mainPath$HTMLSelectedEmailPath")

            ### Replace global tokens in html-body ###
            foreach($gt in $GlobalTokenReplacements.GetEnumerator())
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

            Write-Host "DONE: " $rec.Email
        }
    }
}

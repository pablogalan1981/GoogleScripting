
<#
.SYNOPSIS
    This script will export the GSuite mailbox aliases using GAM tool to a CSV file
    
.DESCRIPTION    
    Before running this script you need to install the GAM tool to export the GSuite mailbox alias info from your Google organization.
    This script will process the data retrieved by GAM tool and will format it to a CSV file. The CSV file will have these columns:  
    resourceId,resourceName,resourceEmail,sourceEmail,destinationEmail
        
    Requirements:
    1. Download and install GAM tool from https://github.com/jay0lee/GAM/releases
.NOTES
	Author		Pablo Galan Sabugo <pablogalan1981@gmail.com> 
	Date		Nov/2018
	Disclaimer: 	This script is provided 'AS IS'. No warrantee is provided either expressed or implied. 
                        BitTitan cannot be held responsible for any misuse of the script.
        Version: 1.1
#>

$csvPath = "C:\scripts"
if ((Test-Path -Path $csvPath) -eq $false) {
    $result = New-Item -Path $csvPath -ItemType directory
    Write-Host -ForegroundColor Green "SUCCESS: Folder $csvPath has been created." 
} 

#Export the resources
Try{
    CMD /C "gam print aliases > $csvPath\MailboxAliases.csv" > $null 2>&1
    Write-Host -ForegroundColor Green "SUCCESS: CSV file $csvPath\MailboxAliases.csv exported from G Suite with GAM tool successfully."
}
Catch [Exception] {
    Write-Host -ForegroundColor Red "ERROR: Failed to export the CSV file 'MailboxAliases.csv' from G Suite with GAM tool."
    Write-Host -ForegroundColor Red $_.Exception.Message
    Exit
}

#Import the CSV file
Try{
    $mailboxAliases = @(Import-CSV "$csvPath\MailboxAliases.csv")
    Write-Host -ForegroundColor Green "SUCCESS: $($mailboxAliases.count) alias have been found."
}
Catch [Exception] {
    Write-Host -ForegroundColor Red "ERROR: Failed to import the CSV file '$csvPath'."
    Write-Host -ForegroundColor Red $_.Exception.Message
    Exit
}

$confirm = (Read-Host -prompt "Are you migrating to the same email addresses?  [Y]es or [N]o")
if($confirm.ToLower() -eq "y") {
    $sameEmailAddresses = $true
}
elseif($confirm.ToLower() -eq "n") {
    $confirm = (Read-Host -prompt "Are you migrating to a different domain?  [Y]es or [N]o")
    if($confirm.ToLower() -eq "y") {
        $differentDomain = $true
        $confirm = (Read-Host -prompt "Do you want to update all alias email addresses to the destination domain?  [Y]es or [N]o")
        if($confirm.ToLower() -eq "y") {
            $updateAliasDomain = $true
            do {
                $destinationDomain = (Read-Host -prompt "Please enter the destination domain")
            }while ($destinationDomain -eq "")
            Write-Host "Destination domain is '$destinationDomain'."
            
            $confirm = (Read-Host -prompt "Are the destination email addresses keeping the same user name part?  [Y]es or [N]o")
            if($confirm.ToLower() -eq "y") {
                $sameUserName = $true
            }
        }
    }    
}

#Declare Output Array
$mailboxAliasesArray = @()
 
foreach ($mailboxAlias in $mailboxAliases) {
    $alias = $mailboxAlias.Alias
    $sourcePrimaryEmail = $mailboxAlias.Target
    $aliasType = $mailboxAlias.TargetType

    
    $aliasLineItem = New-Object PSObject
    $aliasLineItem | Add-Member -MemberType NoteProperty -Name sourcePrimaryEmail -Value $sourcePrimaryEmail
    #Populate destinationPrimaryEmail only when migrating to the same email address
    if ($sameEmailAddresses) {
	    $aliasLineItem | Add-Member -MemberType NoteProperty -Name destinationPrimaryEmail -Value $sourcePrimaryEmail
    }
    else {
        if($sameUserName -and $destinationDomain -ne "") {
            $sourcePrimaryEmailSplit = $sourcePrimaryEmail -split "@"
            $sourceUserName = $sourcePrimaryEmailSplit[0]
            $sourceDomain = $sourcePrimaryEmailSplit[1]
            $newDestinationPrimaryEmail = "$sourceUserName@$destinationDomain" 
            $aliasLineItem | Add-Member -MemberType NoteProperty -Name destinationPrimaryEmail -Value $newDestinationPrimaryEmail
        }
        #If the destination user name and domain change, the destination email will have to be added manually
        else {
            $aliasLineItem | Add-Member -MemberType NoteProperty -Name destinationPrimaryEmail -Value $newDestinationPrimaryEmail
        }
    }
    #Change alias domain if migrating to a new domain and the source domain is not being migrated over
    if($updateAliasDomain) {
        $aliasSplit = $alias -split "@"
        $aliasUserNamer = $aliasSplit[0]
        $aliasDomain = $aliasSplit[1]
        $newAlias = "$aliasUserNamer@$destinationDomain"
        $aliasLineItem | Add-Member -MemberType NoteProperty -Name alias -Value $newAlias
    }
    else {
        $aliasLineItem | Add-Member -MemberType NoteProperty -Name alias -Value $alias
    }
    $aliasLineItem | Add-Member -MemberType NoteProperty -Name aliasType -Value $aliasType

    $mailboxAliasesArray += $aliasLineItem        
} 

#Export mailboxAliasesArray to CSV file
try {
    $mailboxAliasesArray | Export-Csv -Path $csvPath\MailboxAliases.csv -NoTypeInformation -force
    Write-Host -ForegroundColor Green "SUCCESS: CSV file '$csvPath\MailboxAliases.csv' exported successfully."
    if ($sameEmailAddresses) {
        Write-Host -ForegroundColor Yellow "The 'destinationPrimaryEmail' column of the opened CSV file has been populated with the same source email addresses."
        Write-Host -ForegroundColor Yellow "Please review the opened CSV file and once you finish, save it."
    }
    elseif(!$sameEmailAddresses -and $updateAliasDomain -and $destinationDomain -ne "" -and $sameUserName) {
        Write-Host -ForegroundColor Yellow "The 'destinationPrimaryEmail' column of the opened CSV file has been updated with the new domain."
        Write-Host -ForegroundColor Yellow "The 'alias' column of the opened CSV file has been updated with the new domain."
        Write-Host -ForegroundColor Yellow "Please review the opened CSV file and once you finish, save it."
    }
    elseif(!$sameEmailAddresses -and $updateAliasDomain -and $destinationDomain -ne "" -and !$sameUserName) {
        Write-Host -ForegroundColor Yellow "The 'alias' column of the opened CSV file has been updated with the new domain."
        Write-Host -ForegroundColor Yellow "Populate the 'destinationPrimaryEmail' column of the opened CSV file with the destination primary SMTP email addresses."
        Write-Host -ForegroundColor Yellow "Once you finish editing the CSV file, save it."
    }

    elseif (!$sameEmailAddresses -and !$updateAliasDomain) {
        Write-Host -ForegroundColor Yellow "Populate the 'destinationPrimaryEmail' column of the opened CSV file with the destination primary SMTP email addresses."
        Write-Host -ForegroundColor Yellow "The 'alias' column of the opened CSV file contains the source domain. Make sure you have added it to Office 365 before proceeding with the script."
        Write-Host -ForegroundColor Yellow "Once you finish editing the CSV file, save it."
    }
    else {
        Write-Host -ForegroundColor Red "OTRO"
    }
}
catch {
    Write-Host -ForegroundColor Red "ERROR: Failed to import the CSV file '$csvPath\MailboxAliases.csv'."
    Write-Host -ForegroundColor Red $_.Exception.Message
    Exit
}

#Open the CSV file for editing
Start-Process -FilePath $csvPath\MailboxAliases.csv

Write-Host "If you have reviewed and edited the CSV file then press any key to continue." 
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');

#Re-import the edited CSV file
Try{
    $mailboxAliases = Import-CSV "$csvPath\MailboxAliases.csv"
    Write-Host -ForegroundColor Green "SUCCESS: CSV file $csvPath\MailboxAliases.csv re-imported successfully."
}
Catch [Exception] {
    Write-Host -ForegroundColor Red "ERROR: Failed to import the CSV file '$csvPath'. Please save and close the CSV file."
    Write-Host -ForegroundColor Red $_.Exception.Message
    Exit
}

#Add aliases to Office 365 mailboxes

#Prompt for Office 365 global admin Credentials
Write-Host "Connecting to the Office 365 tenant"

$o365Creds = Get-Credential -Message "Enter Office 365 Global Admin credentials"

#Connect to O365
Try
{
    $o365Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $o365Creds -Authentication Basic -AllowRedirection -ErrorAction Stop
    Import-PSSession $o365Session -AllowClobber -ErrorAction Stop
}
Catch
{
    Write-Host -ForegroundColor Red "ERROR: Failed to connect to Office 365."
    Write-Host -ForegroundColor Red $_.Exception.Message
    Exit
}

$userAliasCount = 0
foreach ($mailboxAlias in $mailboxAliases) {
    $alias = $mailboxAlias.Alias
    $sourcePrimaryEmail = $mailboxAlias.sourcePrimaryEmail
    $destinationPrimaryEmail = $mailboxAlias.destinationPrimaryEmail
    $aliasType = $mailboxAlias.aliasType

    if($aliasType -eq "User") {
        $result = Get-Mailbox -Identity $destinationPrimaryEmail  -ErrorAction SilentlyContinue
        if(!$result) {
            Write-Host -ForegroundColor Red "ERROR: Mailbox '$destinationPrimaryEmail' does not exist in Office 365."
        }
        else {

            try {
                $result = Set-Mailbox -Identity $destinationPrimaryEmail -EmailAddresses @{add=$alias} -ErrorAction Stop
                Write-Host -ForegroundColor Green "SUCCESS: Alias '$alias' added to mailbox '$destinationPrimaryEmail'."
                $userAliasCount += 1
            }
            catch {
                Write-Host -ForegroundColor Red "ERROR: Failed to add alias '$alias' to the mailbox '$destinationPrimaryEmail'."
                Write-Host -ForegroundColor Red $_.Exception.Message
                Exit
            }
        }
    }
    elseif($aliasType -eq "Group") {
        
    }
    else {
        Write-Host -ForegroundColor Yellow "WARNING: The alias type '$aliasType' of '$alias' is not recognized. It must be 'User' or 'Group'"
    }
}

if($userAliasCount -ge 1) {
    Write-Host -ForegroundColor Green "SUCCESS: $userAliasCount alias have been added to Office 365 mailboxes."
}

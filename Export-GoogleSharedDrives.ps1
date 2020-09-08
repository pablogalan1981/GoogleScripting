<#

.SYNOPSIS
    Script to export the Google Shared Drives information using GAMADV-XTD tool 

.DESCRIPTION
    Before running this script you need to install the GAM and GAMADV-XTD tools to export the Google Shared Drives info from your Google organization.
    This script will process the data retrieved by GAMADV-XTD tool and will format it to a CSV file. The CSV file will have these columns:    
    TeamDriveId,TeamName,ChannelNames,TeamType,Owners,Members,Roles

    Requirements:
    1. Download and install GAM tool from https://github.com/jay0lee/GAM/releases
    2. Download GAMADV-XTD tool https://github.com/taers232c/GAMADV-XTD/releases 
    3. Upgrade from standard GAM tool to GAMADV-XTD https://github.com/taers232c/GAMADV-XTD/wiki/How-to-Upgrade-from-Standard-GAM#windows 

.NOTES
	Author	     Pablo Galan Sabugo <pablogalan1981@gmail.com> 
	Date         March/2020
	Disclaimer:  This script is provided 'AS IS'. No warrantee is provided either expressed or implied. 
                     BitTitan cannot be held responsible for any misuse of the script.
    Version: 1.1
#>

write-host 
$msg = "####################################################################################`
              EXPORT GOOGLE SHARED DRIVES TO CSV FILE WITH GAMADV TOOL             `
####################################################################################"
Write-Host $msg

$csvPath = "C:\scripts"
if ((Test-Path -Path $csvPath) -eq $false) {
    $result = New-Item -Path $csvPath -ItemType directory
    Write-Host -ForegroundColor Green "SUCCESS: Folder $csvPath has been created." 
} 

#Export the resources with GAM tool from G Suite
Try{
    $TeamDrives = @(CMD /C "gam all users show teamdrives") 
    $TeamDrivesCount = $TeamDrives.Count
    Write-Host -ForegroundColor Green "SUCCESS: $TeamDrivesCount Shared Drives exported from G Suite with GAM tool successfully."
}
Catch [Exception] {
    Write-Host -ForegroundColor Red "ERROR: Failed to export the CSV file 'ResourceMailboxes.csv' from G Suite with GAM tool."
    Write-Host -ForegroundColor Red $_.Exception.Message
    Exit
}

write-host 
$msg = "####################################################################################`
                       ANALYZE EACH SHARED DRIVE              `
####################################################################################"
Write-Host $msg

$teamDriveArray = @()
foreach($TeamDrive in $TeamDrives) {
    $TeamDriveName = ($TeamDrive -split "ID: ")[0] -split "Name: " | ? {$_}
    $TeamDriveId = ($TeamDrive -split "ID: ")[1]

    $lineItem = @{    
        TeamDriveName = $TeamDriveName              
        TeamDriveId = $TeamDriveId    
    }
    $teamDriveArray += New-Object psobject -Property $lineItem
}

$finalTeamDriveArray = @()
foreach($teamDrive in $teamDriveArray) {

    $TeamDriveAcls = @(CMD /C "gam all users show drivefileacl $($teamDrive.TeamDriveId)")

    $membersString = ""
    $ownersString = ""

    for($i=0 ; $i -lt ($TeamDriveAcls.length/10); $i++) {

        $isDeleted = ($TeamDriveAcls[3 + $i*10] -split "deleted: ")[1]

        if($isDeleted) {Continue}

        $role = ($TeamDriveAcls[4 + $i*10] -split "role: ")[1]

        switch ($role) {
            organizer { $owner = ($TeamDriveAcls[3 + $i*10] -split "emailAddress: ")[1]    }
            default { $member = ($TeamDriveAcls[3 + $i*10] -split "emailAddress: ")[1]    }
        }

        if([string]::IsNullOrEmpty($membersString)) {
            $membersString += $member + ";"
        }
        else {
            if($member) {$member=$member.toLower()} else {$member=''}
            if(!($membersString.toLower().contains($member))) {$membersString += $member + ";"}
        } 

        if([string]::IsNullOrEmpty($ownersString)) {
            $ownersString += $owner + ";"
        }
        else { 
            if($owner) {$owner=$owner.toLower()} else {$owner=''}
            if(!($ownersString.toLower().contains($owner))) {$ownersString += $owner + ";"}
        }

        $roleString += $role + ";"
    }

    if($membersString) {$membersString = $membersString.TrimEnd(";")}
    if($ownersString) {$ownersString = $ownersString.TrimEnd(";")} 
    if($roleString) {$roleString = $roleString.TrimEnd(";")} 

    $finalTeamDriveItem = New-Object PSObject
    $finalTeamDriveItem | Add-Member -MemberType NoteProperty -Name TeamDriveId -Value $teamDrive.TeamDriveId 
	$finalTeamDriveItem | Add-Member -MemberType NoteProperty -Name TeamName -Value  $teamDrive.TeamDriveName 
    $finalTeamDriveItem | Add-Member -MemberType NoteProperty -Name ChannelNames -Value ''
    $finalTeamDriveItem | Add-Member -MemberType NoteProperty -Name TeamType -Value "Public"
    $finalTeamDriveItem | Add-Member -MemberType NoteProperty -Name Owners -Value $ownersString
    $finalTeamDriveItem | Add-Member -MemberType NoteProperty -Name Members -Value $membersString
    $finalTeamDriveItem | Add-Member -MemberType NoteProperty -Name Roles -Value $roleString

    $finalTeamDriveArray += $finalTeamDriveItem 
}

write-host 
$msg = "####################################################################################`
                       EXPORT TO CSV FILE             `
####################################################################################"
Write-Host $msg

do {
    try {
        $finalTeamDriveArray | Export-Csv -Path $csvPath\SharedDrives.csv -NoTypeInformation -force -Encoding UTF8 -ErrorAction Stop
        Write-Host -ForegroundColor Green "SUCCESS: CSV file '$csvPath\SharedDrives.csv' exported successfully."

        Break
    }
    catch {
        $msg = "WARNING: Close opened CSV file '$workingDir\SharedDrives.csv'."
        Write-Host -ForegroundColor Yellow $msg
        Write-Host

        Sleep 5
    }
} while ($true)

#Open the CSV file for editing
Start-Process -FilePath $csvPath\SharedDrives.csv

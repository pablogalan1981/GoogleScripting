<#

.SYNOPSIS
    This script will export the GSuite resource mailboxes using GAM tool to a CSV file

.DESCRIPTION    
    Before running this script you need to install the GAM tool to export the GSuite resource mailbox info from your Google organization.
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

#Export the resources with GAM tool from G Suite
Try{
    CMD /C "gam print resources id email Name > $csvPath\ResourceMailboxes.csv"
    Write-Host -ForegroundColor Green "SUCCESS: CSV file $csvPath\ResourceMailboxes.csv exported from G Suite with GAM tool successfully."
}
Catch [Exception] {
    Write-Host -ForegroundColor Red "ERROR: Failed to export the CSV file 'ResourceMailboxes.csv' from G Suite with GAM tool."
    Write-Host -ForegroundColor Red $_.Exception.Message
    Exit
}

#Import the CSV file
Try{
    $resources = @(Import-CSV "$csvPath\ResourceMailboxes.csv")
    Write-Host -ForegroundColor Green "SUCCESS: CSV file $csvPath\ResourceMailboxes.csv imported successfully."
}
Catch [Exception] {
    Write-Host -ForegroundColor Red "ERROR: Failed to import the CSV file '$csvPath\ResourceMailboxes.csv'."
    Write-Host -ForegroundColor Red $_.Exception.Message
    Exit
}

#Ask for the name of the project.
$resourceAdmin = Read-Host "Please enter the account you want to subcribe to all resource mailboxes"

#Declare Output Array
$resourceMailboxesArray = @()
 
foreach ($resource in $resources) {
    $resourceId = $resource.resourceId
    $resourceName = $resource.resourceName
    $resourceEmail = $resource.resourceEmail
    CMD /C "gam user $resourceAdmin add calendar $($resource.resourceEmail) selected true"  

    $resourceLineItem = New-Object PSObject
    $resourceLineItem | Add-Member -MemberType NoteProperty -Name resourceId -Value $resourceId
	$resourceLineItem | Add-Member -MemberType NoteProperty -Name resourceName -Value $resourceName
    $resourceLineItem | Add-Member -MemberType NoteProperty -Name resourceEmail -Value $resourceEmail
    $resourceLineItem | Add-Member -MemberType NoteProperty -Name sourceEmail -Value $resourceAdmin
    $resourceLineItem | Add-Member -MemberType NoteProperty -Name destinationEmail -Value ""

    $resourceMailboxesArray += $resourceLineItem        
} 

#Export resourceMailboxesArray to CSV file
try {
    $resourceMailboxesArray | Export-Csv -Path $csvPath\ResourceMailboxes.csv -NoTypeInformation -force
    Write-Host -ForegroundColor Green "SUCCESS: CSV file '$csvPath\ResourceMailboxes.csv' exported successfully."
    Write-Host -ForegroundColor Yellow "ACTION: Populate the destinationEmail colum of the opened CSV file with the email addresses of the On-Premises Exchange or Exchange Online resource mailboxes."
    Write-Host -ForegroundColor Yellow "        Once you finish editing the CSV file, save it."
}
catch {
    Write-Host -ForegroundColor Red "ERROR: Failed to export the CSV file '$csvPath\ResourceMailboxes.csv'."
    Write-Host -ForegroundColor Red $_.Exception.Message
    Exit
}

#Open the CSV file for editing
Start-Process -FilePath $csvPath\ResourceMailboxes.csv
 

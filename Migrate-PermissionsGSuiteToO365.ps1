<#

.SYNOPSIS

.DESCRIPTION
    Script to     
    1. Generate Gsuite permissions reports
    2. Generate user batches based on FullAccess permissions
    3. Migrate mailbox and folder permissions to O365
	
.NOTES
    Author	    Pablo Galan Sabugo <pablogalan1981@gmail.com> 
    Date	    Nov/2018
    Disclaimer:     This script is provided 'AS IS'. No warrantee is provided either expressed or implied. 
    Version: 1.1
    Change log:
    1.0 - Intitial Draft
#>

#######################################################################################################################
#                                               FUNCTIONS
#######################################################################################################################

### Function to create the working and log directories
Function Create-Working-Directory {    
    param 
    (
        [CmdletBinding()]
        [parameter(Mandatory=$true)] [string]$workingDir,
        [parameter(Mandatory=$true)] [string]$logDir
    )
    if ( !(Test-Path -Path $workingDir)) {
		try {
			$suppressOutput = New-Item -ItemType Directory -Path $workingDir -Force -ErrorAction Stop
            $msg = "SUCCESS: Folder '$($workingDir)' for CSV files has been created."
            Write-Host -ForegroundColor Green $msg
		}
		catch {
            $msg = "ERROR: Failed to create '$workingDir'. Script will abort."
            Write-Host -ForegroundColor Red $msg
            Exit
		}
    }
    if ( !(Test-Path -Path $logDir)) {
        try {
            $suppressOutput = New-Item -ItemType Directory -Path $logDir -Force -ErrorAction Stop      

            $msg = "SUCCESS: Folder '$($logDir)' for log files has been created."
            Write-Host -ForegroundColor Green $msg 
        }
        catch {
            $msg = "ERROR: Failed to create log directory '$($logDir)'. Script will abort."
            Write-Host -ForegroundColor Red $msg
            Exit
        } 
    }
}

### Function to write information to the Log File
Function Log-Write
{
    param
    (
        [Parameter(Mandatory=$true)]    [string]$Message
    )
    $lineItem = "[$(Get-Date -Format "dd-MMM-yyyy HH:mm:ss") | PID:$($pid) | $($env:username) ] " + $Message
	Add-Content -Path $logFile -Value $lineItem
}

###
Function Get-FileName($initialDirectory) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $global:inputFile = $OpenFileDialog.filename

    if($OpenFileDialog.filename -eq "") {
		    # create new import file
	        $inputFileName = "O365Users-import-$((Get-Date).ToString("yyyyMMddHHmmss")).csv"
            $global:inputFile = "C:\Scripts\O365Users-import-$((Get-Date).ToString("yyyyMMddHHmmss")).csv"

		    #$csv = "primarySmtpAddress`r`n"
		    $file = New-Item -Path $initialDirectory -name $inputFileName -ItemType file -force #-value $csv

		    # open file for editing
		    Start-Process excel -FilePath $inputFile

		    do {
			    $confirm = (Read-Host -prompt "Are you done editing the import CSV file?  [Y]es or [N]o")
		        if($confirm -eq "Y") {
			        $importConfirm = $true
		        }

		        if($confirm -eq "N") {
			        $importConfirm = $false
		        }
		    }
		    while(-not $importConfirm)
            
            $msg = "SUCCESS: CSV file '$inputFile' created."
            Write-Host -ForegroundColor Green  $msg
            Log-Write -Message $msg
    }
    else{
        $msg = "INFO: CSF file '$($OpenFileDialog.filename)' selected."
        Write-Host -ForegroundColor Gray  $msg
        Log-Write -Message $msg
    }
}

### Function to create destination EXO PowerShell session
Function Connect-ExchangeOnlineDestination {
    
    #Prompt for destination Office 365 global admin Credentials
    Write-Host "INFO: Connecting to the destination Office 365 tenant"

    if (!($destinationO365Session.State)) {
        try {
            $loginAttempts = 0
            do {
                $loginAttempts++
                # Connect to destination Exchange Online
                $script:destinationO365Creds = Get-Credential -Message "Enter Your Destination Office 365 Admin Credentials"
                if (!($destinationO365Creds)) {
                    $msg = "ERROR: Cancel button or ESC was pressed while asking for Credentials. Script will abort."
                    Write-Host -ForegroundColor Red  $msg
                    Log-Write -Message $msg
                    Exit
                }
                $destinationO365Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $destinationO365Creds -Authentication Basic -AllowRedirection -ErrorAction Stop -WarningAction SilentlyContinue
                $result =Import-PSSession -Session $destinationO365Session -AllowClobber -ErrorAction Stop -WarningAction silentlyContinue -DisableNameChecking
                $msg = "INFO: Remote Exchange PowerShell to destination Office 365 successfully created."
                Log-Write -Message $msg
            }
            until (($loginAttempts -ge 3) -or ($($destinationO365Session.State) -eq "Opened"))

            # Only 3 attempts allowed
            if($loginAttempts -ge 3) {
                $msg = "ERROR: Failed to connect to the destination Office 365. Review your destination Office 365 admin credentials and try again."
                Write-Host $msg -ForegroundColor Red
                Log-Write -Message $msg
                Start-Sleep -Seconds 5
                Exit
            }
        }
        catch {
            $msg = "ERROR: Failed to connect to destination Office 365."
            Write-Host -ForegroundColor Red $msg
            Write-Host -ForegroundColor Red $_.Exception.Message
            Get-PSSession | Remove-PSSession
            Exit
        }        
    } 
    else {
        Get-PSSession | Remove-PSSession
    }

    return $destinationO365Session
}

### Function to query destination email addresses
Function query-EmailAddressMapping {
    do {
        $confirm = (Read-Host -prompt "Are you migrating to the same email addresses?  [Y]es or [N]o")
    } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))

    if($confirm.ToLower() -eq "y") {
        $script:sameEmailAddresses = $true
    }
    elseif($confirm.ToLower() -eq "n") {
        $script:sameEmailAddresses = $false
        do {
            $confirm = (Read-Host -prompt "Are you migrating to a different domain?  [Y]es or [N]o")
        } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))

        if($confirm.ToLower() -eq "y") {
            $script:differentDomain = $true

            do {
                $script:destinationDomain = (Read-Host -prompt "Please enter the destination domain")
            }while ($script:destinationDomain -eq "")
             $msg = "INFO: Destination domain is '$script:destinationDomain'."
             Write-Host $msg

            do{
                $confirm = (Read-Host -prompt "Are the destination email addresses keeping the same user prefix?  [Y]es or [N]o")
            } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))

            if($confirm.ToLower() -eq "y") {
                $script:sameUserName = $true
            }
        }    
    }
}

### Function to check if user exists in destination Office 365
Function check-O365User {
    param 
    (
        [parameter(Mandatory=$true)] [string]$user,
        [parameter(Mandatory=$false)] [string]$userType
    )

    $recipient = Get-Recipient -identity $user -ErrorAction SilentlyContinue

    #$recipientList = @(“UserMailbox”,“SharedMailbox”,“RoomMailbox”,“EquipmentMailbox”,“TeamMailbox”,“GroupMailbox”,“DiscoveryMailbox”,
    #                   “MailContact”,“MailUser”,“MailUniversalDistributionGroup”,“MailUniversalSecurityGroup”,“DynamicDistributionGroup”,“RoomList”,“PublicFolder”,“GuestMailUser”)
    $mailboxList = @(“UserMailbox”,“SharedMailbox”,“RoomMailbox”,“EquipmentMailbox”,“TeamMailbox”,“GroupMailbox”,“DiscoveryMailbox”)
    If ($recipient.RecipientType -in $mailboxList -and $recipient.RecipientypeDetails -ne "DiscoveryMailbox") {
        return $true
    }
    else{
        return $false
    }
}

### Function to create batches
Function Create-UserBatches(){
    param(
        [Parameter(Mandatory=$true)]  [array]$InputPermissions
    )
		
    $data = $InputPermissions
    $hashData = $data | Group primarySmtpAddress -AsHashTable -AsString
	$hashDataByDelegate = $data | Group delegateAddress -AsHashTable -AsString
	$usersWithNoDependents = New-Object System.Collections.ArrayList
    $batch = @{}
    $batchCount = 0
    $hashDataSize = $hashData.Count

    $yyyyMMdd = Get-Date -Format 'yyyyMMdd'

	try{
        #Build ArrayList for users with no dependents
        If($hashDataByDelegate["None"].count -gt 0){
		    $hashDataByDelegate["None"] | %{$_.primarySmtpAddress} | %{[void]$usersWithNoDependents.Add($_)}
	    }	    

        #Identify users with no permissions on them, nor them have perms on another
        If($usersWithNoDependents.count -gt 0){
		    $($usersWithNoDependents) | %{
			    if($hashDataByDelegate.ContainsKey($_)){
				    $usersWithNoDependents.Remove($_)
			    }	
		    }
            
            #Remove users with no dependents from hash Data 
            $usersWithNoDependents | %{$hashData.Remove($_)}

		    #Clean out hashData of users in hash data with no delegates, otherwise they'll get batched
		    foreach($key in $($hashData.keys)){
                    if(($hashData[$key] | select -expandproperty delegateAddress ) -eq "None"){
				    $hashData.Remove($key)
			    }
		    }
	    }
        #Execute batch functions
        If(($hashData.count -ne 0) -or ($usersWithNoDependents.count -ne 0)){
            
            while($hashData.count -ne 0) {
                Find-Associations $hashData
            } 

            Write-Host 
            $msg = "INFO: Generating user batches based on FullAcess permissions"
            Write-Host -ForegroundColor Gray $msg
            Log-Write -Message $msg

            Create-UserBatchFile $batch $usersWithNoDependents   
        }         
    }
    catch {
        $msg = "ERROR: $_"
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg
    }
}

### Function to identify permission associations    
Function Find-Associations($hashData){
    try{
        #"Hash Data Size: $($hashData.count)" 
        $nextInHash = $hashData.Keys | select -first 1
        $batch.Add($nextInHash,$hashData[$nextInHash])
	
	    Do{
		    $checkForMatches = $false
		    foreach($key in $($hashData.keys)){
	            $Script:comparisonCounter++ 
			
			    Write-Progress -Activity "Analyzing associated delegates" -status "Items remaining: $($hashData.Count)" `
    		    -percentComplete (($hashDataSize-$hashData.Count) / $hashDataSize*100)
			
	            #Checks
			    $usersHashData = $($hashData[$key]) | %{$_.primarySmtpAddress}
                $usersBatch = $($batch[$nextInHash]) | %{$_.primarySmtpAddress}
                $delegatesHashData = $($hashData[$key]) | %{$_.delegateAddress} 
			    $delegatesBatch = $($batch[$nextInHash]) | %{$_.delegateAddress}

			    $ifMatchesHashUserToBatchUser = [bool]($usersHashData | ?{$usersBatch -contains $_})
			    $ifMatchesHashDelegToBatchDeleg = [bool]($delegatesHashData | ?{$delegatesBatch -contains $_})
			    $ifMatchesHashUserToBatchDelegate = [bool]($usersHashData | ?{$delegatesBatch -contains $_})
			    $ifMatchesHashDelegToBatchUser = [bool]($delegatesHashData | ?{$usersBatch -contains $_})
			
			    If($ifMatchesHashDelegToBatchDeleg -OR $ifMatchesHashDelegToBatchUser -OR $ifMatchesHashUserToBatchUser -OR $ifMatchesHashUserToBatchDelegate){
	                if(($key -ne $nextInHash)){ 
					    $batch[$nextInHash] += $hashData[$key]
					    $checkForMatches = $true
	                }
	                $hashData.Remove($key)
	            }
	        }
	    } Until ($checkForMatches -eq $false)
        
        return $hashData 
	}
	catch{
        $msg = "ERROR: $_"
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg
    }
}

### Function to create batch file
Function Create-UserBatchFile($batchResults,$usersWithNoDepsResults){
	try {
        
        $userBatchesArray = @()
        
	    foreach($key in $batchResults.keys){
            $batchCount++
            $batchName = "BATCH-$batchCount"

		    $output = New-Object System.Collections.ArrayList
		    $($batch[$key]) | %{$output.add($_.primarySmtpAddress)}
		    $($batch[$key]) | %{$output.add($_.delegateAddress)}
                        
            $output | select -Unique | % { $userBatchLineItem = New-Object PSObject;
                                           $userBatchLineItem | Add-Member -MemberType NoteProperty -Name batchName -Value $batchName;
                                           $userBatchLineItem | Add-Member -MemberType NoteProperty -Name primarySmtpAddress -Value $_;
                                           $userBatchesArray += $userBatchLineItem }                                                     
            
        }
	    If($usersWithNoDepsResults.count -gt 0){
		     $batchCount++
		     foreach($primarySmtpAddress in $usersWithNoDepsResults){
		 	    $batchName = "BATCH-NoFullAccess"  
              
                $userBatchLineItem = New-Object PSObject
                $userBatchLineItem | Add-Member -MemberType NoteProperty -Name batchName -Value $batchName
                $userBatchLineItem | Add-Member -MemberType NoteProperty -Name primarySmtpAddress -Value $primarySmtpAddress
                $userBatchesArray += $userBatchLineItem 
	        }
	    }
         $msg = "SUCCESS: User batches created: $batchCount" 
         Write-host -ForegroundColor Green $msg 
         Log-Write -Message $msg

         $msg = "         INFO: Number of comparisons: $Script:comparisonCounter" 
         Write-host -ForegroundColor Gray $msg 
         Log-Write -Message $msg
    }
    catch{
        $msg = "ERROR: $_"
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg
    }

    $userBatchesArray | Export-Csv -Path $workingDir\GSuiteUserBatches.csv -NoTypeInformation -force

    $msg = "SUCCESS: CSV file '$workingDir\GSuiteUserBatches.csv' processed, exported and open."
    write-Host -ForegroundColor Green $msg
    Log-Write -Message $msg

    #Open the CSV file for editing
    Start-Process -FilePath $workingDir\GSuiteUserBatches.csv

    Write-Host "ACTION: If you have reviewed, edited and saved the CSV file then press any key to continue." 
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
    
} 

#######################################################################################################################
#                                        EXPORT GSUITE PERMISSIONS
#######################################################################################################################
Function Export-GsuitePermissions {
    param 
    (
        [parameter(Mandatory=$true)] [boolean]$skipNonExistingUser,
        [parameter(Mandatory=$true)] [boolean]$processSendAs,
        [parameter(Mandatory=$true)] [boolean]$processFullAccess,
        [parameter(Mandatory=$true)] [boolean]$processCalendars,
        [parameter(Mandatory=$true)] [boolean]$userBatches,
        [parameter(Mandatory=$true)] [boolean]$readCSVfile
    )

    Write-Host
    $msg = "INFO: Exporting all permissions from all mailboxes from GSuite with GAM tool."
    Write-host -ForegroundColor Gray $msg 
    Log-Write -Message $msg

    ########################################################### 
	# Export SendAs permissions	
    ##########################################################  
    if($processSendAs) {
        Write-Host
        Write-Host "INFO: Exporting all SendAs permissions from GSuite with GAM tool."
        Try{
            $gamCommand = "gam all users print sendas > $workingDir\GSuiteSendAs.csv"
            CMD /C "$gamCommand" > $null 2>&1

            $msg = "SUCCESS: SendAs permissions exported from G Suite to CSV file '$workingDir\GSuiteSendAs.csv'."
            Write-Host -ForegroundColor Green $msg
            Log-Write -Message $msg
        }
        Catch [Exception] {
            $msg = "ERROR: Failed to export the CSV file '$workingDir\GSuiteSendAs.csv' from G Suite with GAM tool."
            Write-Host -ForegroundColor Red  $msg
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $msg
            Log-Write -Message $_.Exception.Message
            Exit
        }

        #Import the CSV files
        Try{
            $sendAsPermissions = @(Import-CSV "$workingDir\GSuiteSendAs.csv" | Sort-Object user,displayName | where-Object { $_.PSObject.Properties.Value -ne ""})
            #Write-Host -ForegroundColor Green "SUCCESS: CSV file '$workingDir\GSuiteSendAs.csv' imported."
        }
        Catch [Exception] {
            $msg = "ERROR: Failed to import the CSV file '$workingDir\GSuiteSendAs.csv'."
            Write-Host -ForegroundColor Red  $msg
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $msg
            Log-Write -Message $_.Exception.Message
            Exit
        }

        #Declare Output Arrays
        $script:sendAsPermissionsArray = @()

        $totalUserCount = ($sendAsPermissions | Select user -Unique).count
        $currentUserCount = 0
        $currentUser = $null
        $skipUser = $null
         
        foreach ($sendAsPermission in $sendAsPermissions) {
            ########################################################## 
            # Current SendAsPermission 
            ########################################################## 
            $user = $sendAsPermission.User
            $sendAsName = $sendAsPermission.displayName
            $sendAsEmail = $sendAsPermission.sendAsEmail
            $replyToAddress = $sendAsPermission.replyToAddress
            $isPrimary = $sendAsPermission.isPrimary
            $isDefault = $sendAsPermission.isDefault
            $treatAsAlias = $sendAsPermission.treatAsAlias
            $verificationStatus = $sendAsPermission.verificationStatus
            $signature = $sendAsPermission.signature

            if($currentUser -ne $user) {
                $currentUserCount += 1
                $currentUser = $user
                $userSplit = $user -split("@") 
                $userName = $userSplit[0]

                $msg = "INFO: Processing GSuite user $currentUserCount/$totalUserCount : '$userName' $user."
                Write-host -ForegroundColor Gray $msg 
                Log-Write -Message $msg

                # Skip the user process only if the 3 conditions below are met:
                # 1. Importing permissions into GSuite (not only exporting permissions from O365)
                # 2. The GSuite email addresses do not have different userName and different domain
                # 3. The user does not exist in GSuite

                $result=check-O365User -user $userName 
                if($result){
                    $msg = "      SUCCESS: User '$userName' found in Office 365."
                    Write-Host -ForegroundColor Green $msg
                    Log-Write -Message $msg
                }
                elseif(!$result -and !$skipNonExistingUser -and !$onlyPermissionsReport -and ($script:sameEmailAddresses -or (!$script:sameEmailAddresses -and $script:sameUserName -and $script:destinationDomain -ne ""))) {
                    $msg = "      ERROR: Target mailbox '$user' doest not exist in Office 365."
                    Write-Host -ForegroundColor Red $msg
                    $msg = "      Skipping user processing."
                    Write-Host -ForegroundColor Red $msg
                    Continue
                }
                elseif($skipNonExistingUser){
                    $msg = "      ERROR: Target mailbox '$user' doest not exist in Office 365."
                    Write-Host -ForegroundColor Red $msg
                    $msg = "      Skipping user processing."
                    Write-Host -ForegroundColor Red $msg
                    Continue
                }
            }

            if($skipUser -eq $true){
                Continue
            }
            
            if($isDefault -eq $false) {
       
                $sendAsLineItem = New-Object PSObject

                #Populate destinationPrimaryEmail only when migrating to the same email address
                if ($script:sameEmailAddresses) {
                    $sendAsLineItem | Add-Member -MemberType NoteProperty -Name user -Value $user
                    $sendAsLineItem | Add-Member -MemberType NoteProperty -Name sendAsName -Value $sendAsName
                    $sendAsLineItem | Add-Member -MemberType NoteProperty -Name sendAsEmail -Value $sendAsEmail
                    $sendAsLineItem | Add-Member -MemberType NoteProperty -Name verificationStatus -Value $verificationStatus	    
                }
                else {
                    #If only the destination domain changes and it was entered
                    if($script:sameUserName -and $script:destinationDomain -ne "") {
                        $userSplit = $user -split "@"
                        $userName = $userSplit[0]
                        $userDomain = $userSplit[1]
                        $newUser = "$userName@$script:destinationDomain" 

                        $sendAsEmailSplit = $sendAsEmail -split "@"
                        $sendAsUserName = $sendAsEmailSplit[0]
                        $sendAsDomain = $sendAsEmailSplit[1]
                        $newSendAsEmail = "$sendAsUserName@$script:destinationDomain" 
            
                        $sendAsLineItem | Add-Member -MemberType NoteProperty -Name user -Value $newUser
                        $sendAsLineItem | Add-Member -MemberType NoteProperty -Name sendAsName -Value $sendAsName
                        #if userDomain and sendAsDomain are the same in G Suite, sendAsEmail domain is automatically changed to the new domain
                        if($userDomain -eq $sendAsDomain) {
                            $sendAsLineItem | Add-Member -MemberType NoteProperty -Name sendAsEmail -Value $newSendAsEmail
                        }
                        #if userDomain and sendAsDomain are not the same in G Suite, sendAsEmail domain must be entered manually by the user in the CSV file
                        else {
                            $sendAsLineItem | Add-Member -MemberType NoteProperty -Name sendAsEmail -Value $sendAsEmail
                            $sendAsLineItem | Add-Member -MemberType NoteProperty -Name destinationSendAsEmail -Value ""
                        }                
                        $sendAsLineItem | Add-Member -MemberType NoteProperty -Name verificationStatus -Value $verificationStatus	
                    }
                    #If the destination user name changes or the new domain was not entered, the user and sendAsEmail emails must be entered manually by the user in the CSV file
                    else {
                        $sendAsLineItem | Add-Member -MemberType NoteProperty -Name user -Value $user
                        $sendAsLineItem | Add-Member -MemberType NoteProperty -Name destinationUser -Value ""
                        $sendAsLineItem | Add-Member -MemberType NoteProperty -Name sendAsName -Value $sendAsName
                        $sendAsLineItem | Add-Member -MemberType NoteProperty -Name sendAsEmail -Value $sendAsEmail
                        $sendAsLineItem | Add-Member -MemberType NoteProperty -Name destinationSendAsEmail -Value ""
                        $sendAsLineItem | Add-Member -MemberType NoteProperty -Name verificationStatus -Value $verificationStatus	
                    }
                }

                $script:sendAsPermissionsArray += $sendAsLineItem
            }    
        } 
    }
       
    ########################################################### 
	# Export FullAccess permissions	
    ########################################################## 
    if($processFullAccess) {
        Write-Host
        Write-Host "INFO: Exporting all FullAccess permissions from GSuite with GAM tool."  
        Try{
            $gamCommand = "gam all users print delegates > $workingDir\GSuiteFullAccess.csv"
            CMD /C "$gamCommand" > $null 2>&1

            $msg = "SUCCESS: FullAccess permissions exported from G Suite to CSV file '$workingDir\GSuiteFullAccess.csv'."
            Write-Host -ForegroundColor Green $msg
            Log-Write -Message $msg
        }
        Catch [Exception] {             
            $msg = "ERROR: Failed to export the CSV file '$workingDir\GSuiteFullAccess.csv' from G Suite with GAM tool."
            Write-Host -ForegroundColor Red  $msg
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $msg
            Log-Write -Message $_.Exception.Message
            Exit
        }

        #Import the CSV files
        Try{
            $fullAccessPermissions = @(Import-CSV "$workingDir\GSuiteFullAccess.csv" | Sort-Object user | where-Object { $_.PSObject.Properties.Value -ne ""})
            #Write-Host -ForegroundColor Green "SUCCESS: CSV file '$workingDir\GSuiteFullAccess.csv' imported."
        }
        Catch [Exception] {
            $msg = "ERROR: Failed to import the CSV file '$workingDir\GSuiteFullAccess.csv'."
            Write-Host -ForegroundColor Red  $msg
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $msg
            Log-Write -Message $_.Exception.Message
            Exit
        }

        #Declare Output Arrays
        $script:fullAccessPermissionsArray = @()
        $userBatchesArray = @()

        $totalUserCount = ($fullAccessPermissions | Select user -Unique).count
        $currentUserCount = 0
        $currentUser = $null
 
        foreach ($fullAccessPermission in $fullAccessPermissions) {
            $user = $fullAccessPermission.User
            $delegateName = $fullAccessPermission.delegateName
            $delegateAddress = $fullAccessPermission.delegateAddress
            $delegationStatus = $fullAccessPermission.delegationStatus
       
            if($currentUser -ne $user) {
                $currentUserCount += 1
                $currentUser = $user
                $userSplit = $user -split("@") 
                $userName = $userSplit[0]

                $msg = "INFO: Processing user $currentUserCount/$totalUserCount : '$userName' $user."
                Write-host -ForegroundColor Gray $msg 
                Log-Write -Message $msg

                # Skip the user process only if the 3 conditions below are met:
                # 1. Importing permissions into GSuite (not only exporting permissions from O365)
                # 2. The GSuite email addresses do not have different userName and different domain
                # 3. The user does not exist in GSuite

                $result=check-O365User -user $userName
                if($result){
                    $msg = "      SUCCESS: User '$userName' found in Office 365."
                    Write-Host -ForegroundColor Green $msg
                    Log-Write -Message $msg
                }
                elseif(!$result -and !$skipNonExistingUser -and !$onlyPermissionsReport -and ($script:sameEmailAddresses -or (!$script:sameEmailAddresses -and $script:sameUserName -and $script:destinationDomain -ne ""))) {
                    $msg = "      ERROR: Target mailbox '$user' doest not exist in Office 365."
                    Write-Host -ForegroundColor Red $msg
                    $msg = "      Skipping user processing."
                    Write-Host -ForegroundColor Red $msg
                    Continue
                }
                elseif($skipNonExistingUser){
                    $msg = "      ERROR: Target mailbox '$user' doest not exist in Office 365."
                    Write-Host -ForegroundColor Red $msg
                    $msg = "      Skipping user processing."
                    Write-Host -ForegroundColor Red $msg
                    Continue
                }
            }

            if($skipUser -eq $true){
                Continue
            }

            $fullAccessLineItem = New-Object PSObject

            #Populate destinationPrimaryEmail only when migrating to the same email address
            if ($script:sameEmailAddresses) {
                $fullAccessLineItem | Add-Member -MemberType NoteProperty -Name user -Value $user
                $fullAccessLineItem | Add-Member -MemberType NoteProperty -Name delegateName -Value $delegateName
                $fullAccessLineItem | Add-Member -MemberType NoteProperty -Name delegateAddress -Value $delegateAddress
                $fullAccessLineItem | Add-Member -MemberType NoteProperty -Name delegationStatus -Value $delegationStatus
                $userBatchesArray += $fullAccessLineItem	    
            }
            else {
                #If only the destination domain changes and it was entered
                if($script:sameUserName -and $script:destinationDomain -ne "") {
                    $userSplit = $user -split "@"
                    $userName = $userSplit[0]
                    $userDomain = $userSplit[1]
                    $newUser = "$userName@$script:destinationDomain" 

                    $fullAccessEmailSplit = $delegateAddress -split "@"
                    $fullAccessUserName = $fullAccessEmailSplit[0]
                    $fullAccessDomain = $fullAccessEmailSplit[1]
                    $newFullAccessEmail = "$fullAccessUserName@$script:destinationDomain" 
            
                    $fullAccessLineItem | Add-Member -MemberType NoteProperty -Name user -Value $newUser
                    #if userDomain and fullAccessDomain are the same in G Suite, delegate domain is automatically changed to the new domain
                    if($userDomain -eq $fullAccessDomain) {
                        $fullAccessLineItem | Add-Member -MemberType NoteProperty -Name delegateAddress -Value $newFullAccessEmail
                    }
                    #if userDomain and fullAccessDomain are not the same in G Suite, delegate domain must be entered manually by the user in the CSV file
                    else {
                        $fullAccessLineItem | Add-Member -MemberType NoteProperty -Name delegateAddress -Value $delegateAddress
                        $fullAccessLineItem | Add-Member -MemberType NoteProperty -Name destinationDelegateAddress -Value ""
                    }
                    $fullAccessLineItem | Add-Member -MemberType NoteProperty -Name delegateName -Value $delegateName
                    $fullAccessLineItem | Add-Member -MemberType NoteProperty -Name delegationStatus -Value $delegationStatus	
                }
                #If the destination user name changes or the new domain was not entered, the user and delegate emails must be entered manually by the user in the CSV file
                else {
                    $fullAccessLineItem | Add-Member -MemberType NoteProperty -Name user -Value $user
                    $fullAccessLineItem | Add-Member -MemberType NoteProperty -Name destinationUser -Value ""
                    $fullAccessLineItem | Add-Member -MemberType NoteProperty -Name delegateName -Value $delegateName
                    $fullAccessLineItem | Add-Member -MemberType NoteProperty -Name delegateAddress -Value $delegateAddress
                    $fullAccessLineItem | Add-Member -MemberType NoteProperty -Name destinationDelegateAddress -Value ""
                    $fullAccessLineItem | Add-Member -MemberType NoteProperty -Name delegationStatus -Value $delegationStatus	
                }
            }
            $script:fullAccessPermissionsArray += $fullAccessLineItem 
            $userBatchesArray += $fullAccessLineItem        
        }

        if($createUserBatches -eq $true -and $userBatchesArray -ne $null){
            Create-UserBatches -InputPermissions $userBatchesArray
        }
    }
    
    ########################################################### 
	# Export calendar permissions	
    ########################################################## 
    if($processCalendars) {
        Write-Host
        Write-Host "INFO: Exporting all Calendar permissions from GSuite with GAM tool." 
        Try{
            $gamCommand = "gam print users | gam csv - gam calendar ~primaryEmail  showacl > $workingDir\GSuiteCalendarACL.csv"
            CMD /C "$gamCommand" > $null 2>&1

            $msg = "SUCCESS: Calendar permissions exported from G Suite to CSV file '$workingDir\GSuiteCalendarACL.csv'."
            Write-Host -ForegroundColor Green $msg
            Log-Write -Message $msg
        }
        Catch [Exception] {
            $msg = "ERROR: Failed to export the CSV file '$workingDir\GSuiteCalendarACL.csv' from G Suite with GAM tool."
            Write-Host -ForegroundColor Red  $msg
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $msg
            Log-Write -Message $_.Exception.Message
            Exit
        }

        #Import the CSV files
        Try{
            $calendarPermissions = @(Import-CSV "$workingDir\GSuiteCalendarACL.csv" -Header Calendar,Scope,Role | Sort-Object calendar | where-Object { $_.PSObject.Properties.Value -ne ""})
            #Write-Host -ForegroundColor Green "SUCCESS: CSV file '$workingDir\GSuiteCalendarACL.csv' imported."
        }
        Catch [Exception] {
            $msg = "ERROR: Failed to import the CSV file '$workingDir\GSuiteCalendarACL.csv'."
            Write-Host -ForegroundColor Red  $msg
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $msg
            Log-Write -Message $_.Exception.Message
            Exit
        }

        #Declare Output Arrays
        $script:calendarPermissionsArray = @()

        $totalUserCount = ($calendarPermissions | Select calendar -Unique).count
        $currentUserCount = 0
        $currentUser = $null
 
        foreach ($calendarPermission in $calendarPermissions) {

           $calendar = $calendarPermission.Calendar.replace("Calendar: ","") 
           $scopeType = $calendarPermission.Scope.replace("ACL: (Scope: ","").split(":")[0]
           $scope = $calendarPermission.Scope.replace("ACL: (Scope: ","").split(":")[1]
           $googleRole = $calendarPermission.role.replace("Role: ","").split(")")[0]

           if($currentUser -ne $calendar) {
                $currentUserCount += 1
                $currentUser = $calendar
                $userSplit = $calendar -split("@") 
                $userName = $userSplit[0]

                $msg = "INFO: Processing user $currentUserCount/$totalUserCount : '$userName' $calendar."
                Write-host -ForegroundColor Gray $msg 
                Log-Write -Message $msg

                $result=check-O365User -user $userName
                # Not skip user if it exists in O365 
                if($result){
                    $msg = "      SUCCESS: User '$userName' found in Office 365."
                    Write-Host -ForegroundColor Green $msg
                    Log-Write -Message $msg
                    $skipUser=$false
                }
                # Not skip user if it does not exists in O365 AND $skipNonExistingUser=$false
                elseif(!$result -and !$skipNonExistingUser){
                    $msg = "      ERROR: Target mailbox '$userName' doest not exist in Office 365."
                    Write-Host -ForegroundColor Red $msg
                    $msg = "      Skipping user processing."
                    Write-Host -ForegroundColor Red $msg
                    $skipUser=$false
                }
                # Skip user if it does not exists in O365 AND $skipNonExistingUser=$true
                elseif(!$result -and $skipNonExistingUser){
                    $msg = "      ERROR: Target mailbox '$userName' doest not exist in Office 365."
                    Write-Host -ForegroundColor Red $msg
                    $msg = "      Skipping user processing."
                    Write-Host -ForegroundColor Red $msg
                    $skipUser=$true
                    Continue
                }
                # Skip user if $skipNonExistingUser=$false AND importing permissions into O365 (not only exporting permissions from GSuite) 
                # AND it does not exists in O365
                #    with the same email addresses OR
                #    with differnt email addresses BUT the same userName with destinationDomain provided
                # Not Skip it does not exists in O365 BUT with different userName and destinationDOmain not provided
                elseif(!$result -and !$skipNonExistingUser -and !$onlyPermissionsReport -and ($script:sameEmailAddresses -or (!$script:sameEmailAddresses -and $script:sameUserName -and $script:destinationDomain -ne ""))) {
                    $msg = "      ERROR: Target mailbox '$userName' doest not exist in Office 365."
                    Write-Host -ForegroundColor Red $msg
                    $msg = "      Skipping user processing."
                    Write-Host -ForegroundColor Red $msg
                    $skipUser=$true
                    Continue
                }
            }

            if($skipUser -eq $true){
                Continue
            }
             
            ################################################################
            #G Suite roles
            ################################################################
            #None            Provides no access.
            #FreeBusyReader  Provides read access to free/busy information.
            #Reader          Provides read access to the calendar. Private events will appear to users with reader access, but event details will be hidden.
            #Writer          Provides read and write access to the calendar. Private events will appear to users with writer access, and event details will be visible.
            #Owner           Provides ownership of the calendar. This role has all of the permissions of the writer role with the additional ability to see and manipulate ACLs.

            ################################################################
            #Office 365 Calendar permission levels
            ################################################################
            #None             User will be unable to view any information (including free/busy times).
            #LimitedDetails   Allows someone to view blocks of time as Free, Busy, Tentative, Away.
            #AvailabilityOnly Allows someone to view your Subject and Location. Events set to private will only display as Private Appointment.
            #Contributor      Provides the ability to view free/busy information and create new events.
            #Reviewer         Allows someone to view your Subject, Location, Attendees, and Description. However, any event you mark as private displays simply as Private Appointment.
            #NonEditingAuthor Provides the ability to view full details of all events (accept private ones), create new events, and delete events they have created.
            #Author           In addition to permissions granted via "Nonediting Authour", the user will also me able to edit events they have created.
            #PublishingAuthor In addition to permissions granted via "Authour", the user will also me able to create sub-folders (these are calendar groups or secondary calendars).
            #Editor           Provides read/write/modify access to the calendar (accept private events).
            #PublishingEditor In addition to "Editor" permissions, the user will also be able to create sub-folders (these are calendar groups or secondary calendars).
            #Owner            In addition to "Editor" permissions, a delegate can also be selected to receive calendar notifications/requests/invitations. By default, 'Delegates' cannot view/modify events set to Private. 
            #                 You do have the option to grant the delegate the ability to view (full details) Private events.

           switch ( $googleRole ){
                owner          { $o365Role = "Owner"            }
                writer         { $o365Role = "Editor"           }
                reader         { $o365Role = "Reviewer"         }
                freeBusyReader { $o365Role = "AvailabilityOnly" }
                none           { $o365Role = "None"             }
            }
        
            if($scopeType -eq "user" -and $calendar -ne $scope) {

                $calendarLineItem = New-Object PSObject

                #Populate destinationPrimaryEmail only when migrating to the same email address
                if ($script:sameEmailAddresses) {
                    $calendarLineItem | Add-Member -MemberType NoteProperty -Name calendar -Value $calendar
                    $calendarLineItem | Add-Member -MemberType NoteProperty -Name scopeType -Value $scopeType
                    $calendarLineItem | Add-Member -MemberType NoteProperty -Name scope -Value $scope
                    $calendarLineItem | Add-Member -MemberType NoteProperty -Name googleRole -Value $googleRole	   
                    $calendarLineItem | Add-Member -MemberType NoteProperty -Name o365Role -Value $o365Role	  
                }
                else {
                    #If only the destination domain changes and it was entered
                    if($script:sameUserName -and $script:destinationDomain -ne "") {
                        $calendarSplit = $calendar -split "@"
                        $calendarUserName = $calendarSplit[0]
                        $calendarDomain = $calendarSplit[1]
                        $destinationCalendar = "$calendarUserName@$script:destinationDomain" 
                    
                        $scopeSplit = $scope -split "@"
                        $scopeUserName = $scopeSplit[0]
                        $scopeDomain = $scopeSplit[1]
                        $destinationScope = "$scopeUserName@$script:destinationDomain" 
                                               
                        $calendarLineItem | Add-Member -MemberType NoteProperty -Name calendar -Value $destinationCalendar
                        $calendarLineItem | Add-Member -MemberType NoteProperty -Name scopeType -Value $scopeType
                        #if userDomain and calendarDomain are the same in G Suite, delegate domain is automatically changed to the new domain
                        if($calendarDomain -eq $scopeDomain) {
                            $calendarLineItem | Add-Member -MemberType NoteProperty -Name scope -Value $destinationScope
                        }
                        #if userDomain and calendarDomain are not the same in G Suite, delegate domain must be entered manually by the user in the CSV file
                        else {
                            $calendarLineItem | Add-Member -MemberType NoteProperty -Name scope -Value $scope
                            $calendarLineItem | Add-Member -MemberType NoteProperty -Name destinationScope -Value ""
                        }
                        $calendarLineItem | Add-Member -MemberType NoteProperty -Name googleRole -Value $googleRole	   
                        $calendarLineItem | Add-Member -MemberType NoteProperty -Name o365Role -Value $o365Role	  
                    }
                    #If the destination user name changes or the new domain was not entered, the user and delegate emails must be entered manually by the user in the CSV file
                    else {
                        $calendarLineItem | Add-Member -MemberType NoteProperty -Name calendar -Value $calendar
                        $calendarLineItem | Add-Member -MemberType NoteProperty -Name destinationCalendar -Value ""
                        $calendarLineItem | Add-Member -MemberType NoteProperty -Name scopeType -Value $scopeType
                        $calendarLineItem | Add-Member -MemberType NoteProperty -Name scope -Value $scope
                        $calendarLineItem | Add-Member -MemberType NoteProperty -Name destinationScope -Value ""
                        $calendarLineItem | Add-Member -MemberType NoteProperty -Name googleRole -Value $googleRole	   
                        $calendarLineItem | Add-Member -MemberType NoteProperty -Name o365Role -Value $o365Role	  
                    }
                }
                $script:calendarPermissionsArray += $calendarLineItem         
            }     
        } 
    }
}

#######################################################################################################################
#                                        IMPORT SENDAS PERMISSIONS INTO O365
#######################################################################################################################
Function Process-SendAsPermissions {

    Write-Host "INFO: Exporting SendAs permissions to CSV file."
    
    #Export sendAsPermissionsArray to CSV file
    try {
        if($onlyPermissionsReport) {
            $script:sendAsPermissionsArray | Export-Csv -Path $workingDir\GSuiteSendAsReport.csv -NoTypeInformation -force
            Write-Host -ForegroundColor Green "SUCCESS: CSV file '$workingDir\GSuiteSendAsReport.csv' processed, exported and open."
        }
        else {
            $script:sendAsPermissionsArray | Export-Csv -Path $workingDir\GSuiteSendAs.csv -NoTypeInformation -force
            Write-Host -ForegroundColor Green "SUCCESS: CSV file '$workingDir\GSuiteSendAs.csv' processed, exported and open."
        } 

        if ($script:sameEmailAddresses) {
            Write-Host -ForegroundColor Yellow "         ACTION:  Please review the opened CSV file and once you finish, save it."
        }
        elseif(!$script:sameEmailAddresses -and $script:sameUserName -and $script:destinationDomain -ne "" -and $userDomain -eq $sendAsDomain -and $onlyPermissionsReport -eq $false) {
            Write-Host -ForegroundColor Yellow "         WARNING: The 'user' column of the opened CSV file has been updated with the new domain."
            Write-Host -ForegroundColor Yellow "         WARNING: The 'sendAsEmail' column of the opened CSV file has been updated with the new domain."
            Write-Host -ForegroundColor Yellow "         ACTION:  Please review the opened CSV file and once you finish, save it."
        }
        elseif(!$script:sameEmailAddresses -and $script:sameUserName -and $script:destinationDomain -ne ""  -and $userDomain -ne $sendAsDomain -and $onlyPermissionsReport -eq $false) {
            Write-Host -ForegroundColor Yellow "         WARNING: The 'user' column of the opened CSV file has been updated with the new domain."
            Write-Host -ForegroundColor Yellow "         ACTION:  Populate the 'destinationSendAsEmail' column of the opened CSV file with the destination SendAs email."
            Write-Host -ForegroundColor Yellow "         ACTION:  Please review the opened CSV file and once you finish, save it."
        }
        elseif(!$script:sameEmailAddresses -and !$script:sameUserName -and $script:destinationDomain -ne "" -and $onlyPermissionsReport -eq $false) {
            Write-Host -ForegroundColor Yellow "         ACTION:  Populate the 'destinationUser' column of the opened CSV file with the destination user email."
            Write-Host -ForegroundColor Yellow "         ACTION:  Populate the 'destinationSendAsEmail' column of the opened CSV file with the destination SendAs email."
            Write-Host -ForegroundColor Yellow "         ACTION:  Once you finish editing the CSV file, save it."
        }
        elseif ($onlyPermissionsReport -eq $false) {
        }
    }
    catch {
        if($onlyPermissionsReport) {
            $msg = "ERROR: Failed to import the CSV file '$workingDir\GSuiteSendAsReport.csv'."
        }
        else {
            $msg = "ERROR: Failed to import the CSV file '$workingDir\GSuiteSendAs.csv'."
        } 
        Write-Host -ForegroundColor Red  $msg
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $msg
        Log-Write -Message $_.Exception.Message
        Exit
    }

    #Open the CSV file for editing
    if($onlyPermissionsReport) {
        Start-Process -FilePath $workingDir\GSuiteSendAsReport.csv
    }
    else {
        Start-Process -FilePath $workingDir\GSuiteSendAs.csv
    }  

    #If the script must generate GSuite permissions report and also migrate them to O365
    if(!$onlyPermissionsReport) {
        Write-Host "ACTION: If you have reviewed, edited and saved the CSV file then press any key to continue." 
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');

        #Re-import the edited CSV file
        Try{
            $sendAsPermissions = @(Import-CSV "$workingDir\GSuiteSendAs.csv" | where-Object { $_.PSObject.Properties.Value -ne ""})
            #Write-Host -ForegroundColor Green "SUCCESS: CSV file '$workingDir\GSuiteSendAs.csv' re-imported."
        }
        Catch [Exception] {
            $msg = "ERROR: Failed to import the CSV file '$workingDir\GSuiteSendAs.csv'. Please save and close the CSV file."
            Write-Host -ForegroundColor Red  $msg
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $msg
            Log-Write -Message $_.Exception.Message
            Exit
        }

        $totalSendAsPermissionsExport = $sendAsPermissions.Count
        $sendAsPermissionsCount = 0
        $currentSendAsPermission = 0

        Write-Host  
        Write-Host "INFO: Importing SendAs permissions into Office 365."  

        Foreach($sendAsPermission in $sendAsPermissions){
            #Current SendAs Permission
            $currentSendAsPermission += 1
            if($sendAsPermission.destinationUser) {
                $targetMailbox = $sendAsPermission.destinationUser
            }
            else {
                $targetMailbox = $sendAsPermission.user
            }

            if($sendAsPermission.destinationSendAsEmail) {
                $sendAsEmail = $sendAsPermission.destinationSendAsEmail
            }
            else {
                $sendAsEmail = $sendAsPermission.sendAsEmail
            }
            $sendAsName = $sendAsPermission.sendAsName
            $verificationStatus = $sendAsPermission.verificationStatus
                
            $msg = "INFO: Processing SendAs permission $currentSendAsPermission/$totalSendAsPermissionsExport : TargetMailbox $targetMailbox SendAsEmail $sendAsEmail."
            Write-host -ForegroundColor Gray $msg
            Log-Write -Message $msg
            
            #Verify if target mailbox exists    
            $recipient = check-O365User -user $targetMailbox

            If ($recipient -eq $true) {

                #Verify if sendAsEmail exists
                $recipient = check-O365User -user $sendAsEmail 

                If($recipient -eq $true) {
                    if($verificationStatus -ne "pending") {
                        try {
                            $result = Get-RecipientPermission $targetMailbox -Trustee $sendAsEmail -AccessRights "SendAs"
                            if (!$result) {
                                $result=Add-RecipientPermission $targetMailbox -Trustee $sendAsEmail -AccessRights "SendAs" -Confirm:$false 

                                $msg = "      SUCCESS: SendAs permission applied."
                                Write-Host -ForegroundColor Green $msg
                                Log-Write -Message $msg
                                $sendAsPermissionsCount += 1
                            }
                            else {
                                $msg = "      WARNING: SendAs permission already exists in Office 365."
                                Write-Host -ForegroundColor Yellow $msg 
                                Log-Write -Message $msg
                            }
                        }
                        catch {
                            $msg = "      ERROR: Failed to apply SendAs permission."
                            Write-Host -ForegroundColor Red  $msg
                            Write-Host -ForegroundColor Red $_.Exception.Message
                            Log-Write -Message $msg
                            Log-Write -Message $_.Exception.Message
                        }
                    }
                    else {
                        $msg = "      WARNING: The verification status for SendAs permission is still in pending in GSuite." 
                        Write-Host -ForegroundColor Yellow $msg 
                        Log-Write -Message $msg
                        $msg = "      WARNING: SendAs permission won't be re-applied."   
                        Write-Host -ForegroundColor Yellow $msg 
                        Log-Write -Message $msg                 
                    }
                }
                else {
                    $msg = "      ERROR: SendAsEmail '$sendAsEmail' doest not exist in Office 365. SendAs permission skipped."
                    Write-Host -ForegroundColor Red  $msg
                    Log-Write -Message $msg
                }        
            } 
            else{
                $msg =  "      ERROR: Target mailbox '$targetMailbox' doest not exist in Office 365."
                Write-Host -ForegroundColor Red  $msg
                Log-Write -Message $msg
            }  


        }

        if($sendAsPermissionsCount -ge 2) {
            $msg = "SUCCESS: $sendAsPermissionsCount SendAs permissions out of $totalSendAsPermissionsExport have been applied to Office 365 mailboxes."
            Write-Host -ForegroundColor Green $msg
            Log-Write -Message $msg
        }elseif ($sendAsPermissionsCount -eq 1) {
            $msg = "SUCCESS: 1 SendAs permission out of $totalSendAsPermissionsExport has been applied to Office 365 mailboxes."
            Write-Host -ForegroundColor Green $msg
            Log-Write -Message $msg
        }
    }
}

#######################################################################################################################
#                                        IMPORT FULLACCESS PERMISSIONS INTO O365
#######################################################################################################################
Function Process-FullAccessPermissions {

    Write-Host "INFO: Exporting FullAccess permissions to CSV file."

    #Export fullAccessPermissionsArray to CSV file
    try {
        if($onlyPermissionsReport) {
            $script:fullAccessPermissionsArray | Export-Csv -Path $workingDir\GSuiteFullAccessReport.csv -NoTypeInformation -force
            Write-Host -ForegroundColor Green "SUCCESS: CSV file '$workingDir\GSuiteFullAccessReport.csv' processed, exported and open."
        }
        else {
            $script:fullAccessPermissionsArray | Export-Csv -Path $workingDir\GSuiteFullAccess.csv -NoTypeInformation -force
            Write-Host -ForegroundColor Green "SUCCESS: CSV file '$workingDir\GSuiteFullAccess.csv' processed, exported and open."
        }   
        if ($script:sameEmailAddresses) {
            Write-Host -ForegroundColor Yellow "         ACTION:  Please review the opened CSV file and once you finish, save it."
        }
        elseif(!$script:sameEmailAddresses -and $script:sameUserName -and $script:destinationDomain -ne "" -and $userDomain -eq $fullAccessDomain -and $onlyPermissionsReport -eq $false) {
            Write-Host -ForegroundColor Yellow "         WARNING: The 'user' column of the opened CSV file has been updated with the new domain."
            Write-Host -ForegroundColor Yellow "         WARNING: The 'delegateAddress' column of the opened CSV file has been updated with the new domain."
            Write-Host -ForegroundColor Yellow "         ACTION:  Please review the opened CSV file and once you finish, save it."
        }
        elseif(!$script:sameEmailAddresses -and $script:sameUserName -and $script:destinationDomain -ne ""  -and $userDomain -ne $fullAccessDomain -and $onlyPermissionsReport -eq $false) {
            Write-Host -ForegroundColor Yellow "         WARNING: The 'user' column of the opened CSV file has been updated with the new domain."
            Write-Host -ForegroundColor Yellow "         ACTION:  Populate the 'destinationDelegateAddress' column of the opened CSV file with the destination FullAccess email."
            Write-Host -ForegroundColor Yellow "         ACTION:  Please review the opened CSV file and once you finish, save it."
        }
        elseif(!$script:sameEmailAddresses -and !$script:sameUserName -and $script:destinationDomain -ne "" -and $onlyPermissionsReport -eq $false) {
            Write-Host -ForegroundColor Yellow "         WARNING: Populate the 'destinationUser' column of the opened CSV file with the destination user email."
            Write-Host -ForegroundColor Yellow "         WARNING: Populate the 'destinationDelegateAddress' column of the opened CSV file with the destination FullAccess email."
            Write-Host -ForegroundColor Yellow "         ACTION:  Once you finish editing the CSV file, save it."
        }
        elseif($onlyPermissionsReport -eq $false) {
        }
    }
    catch {
        if($onlyPermissionsReport) {
            $msg = "ERROR: Failed to import the CSV file '$workingDir\GSuiteFullAccessReport.csv'."
        }
        else {
            $msg = "ERROR: Failed to import the CSV file '$workingDir\GSuiteFullAccess.csv'."
        } 
        Write-Host -ForegroundColor Red  $msg
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $msg
        Log-Write -Message $_.Exception.Message
        Exit
    }

    #Open the CSV file for editing

    if($onlyPermissionsReport) {
        Start-Process -FilePath $workingDir\GSuiteFullAccessReport.csv
    }
    else {
        Start-Process -FilePath $workingDir\GSuiteFullAccess.csv
    } 

    #If the script must generate GSuite permissions report and also migrate them to O365
    if(!$onlyPermissionsReport) {
        Write-Host "ACTION: If you have reviewed, edited and saved the CSV file then press any key to continue." 
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');

        #Re-import the edited CSV file
        Try{
            $fullAccessPermissions = @(Import-CSV "$workingDir\GSuiteFullAccess.csv" | where-Object { $_.PSObject.Properties.Value -ne ""})
            #Write-Host -ForegroundColor Green "SUCCESS: CSV file '$workingDir\GSuiteFullAccess.csv' re-imported."
        }
        Catch [Exception] {
            $msg = "ERROR: Failed to import the CSV file '$workingDir\GSuiteFullAccess.csv'. Please save and close the CSV file."
            Write-Host -ForegroundColor Red  $msg
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $msg
            Log-Write -Message $_.Exception.Message
            Exit
        }

        $confirm = (Read-Host -prompt "Do you want to enable the auto-mapping feature in Microsoft Outlook that uses Autodiscover?  [Y]es or [N]o")
        if($confirm.ToLower() -eq "y") {
            $autoMapping = $true
        }
        $totalFullAccessPermissionsExport = $fullAccessPermissions.count
        $FullAccessPermissionsCount = 0
        $currentFullAccessPermission = 0

        Write-Host 
        Write-Host "INFO: Importing FullAccess permissions into Office 365."  

        Foreach($FullAccessPermission in $FullAccessPermissions){
            $currentFullAccessPermission += 1
            if($FullAccessPermission.destinationUser) {
                $targetMailbox = $FullAccessPermission.destinationUser
            }
            else {
                $targetMailbox = $FullAccessPermission.user
            }

            if($FullAccessPermission.destinationDelegateAddress) {
                $delegate = $FullAccessPermission.destinationDelegateAddress
            }
            else {
                $delegate = $FullAccessPermission.delegateAddress
            }

            $delegateName = $FullAccessPermission.delegateName
            $delegationStatus = $FullAccessPermission.delegationStatus

            $msg = "INFO: Processing FullAccess permission $currentFullAccessPermission/$totalFullAccessPermissionsExport : TargetMailbox $targetMailbox Delegate $delegate."
            Write-Host -ForegroundColor Gray $msg
            Log-Write -Message $msg
             
            #Verify if target mailbox exists    
            $recipient = check-O365User -user $targetMailbox

            If ($recipient -eq $true) {

                #Verify if delegate exists
                $recipient = check-O365User -user $delegate 

                If($recipient -eq $true) {
                    if($delegationStatus -ne "pending") {

                        $result = Get-MailboxPermission -identity $targetMailbox -User $delegate 

                        if($result.AccessRights -eq "FullAccess") {
                            $msg = "      WARNING: FullAccess permission already exists in Office 365."
                            Write-Host -ForegroundColor Yellow $msg 
                            Log-Write -Message $msg                      
                        }
                        else {
                            try {
                                If($autoMapping -eq $true) {
                                    $result = Add-MailboxPermission -identity $targetMailbox -User $delegate -automapping $true -AccessRights FullAccess -InheritanceType All -ErrorAction Stop
                                }
                                else {
                                    $result = Add-MailboxPermission -identity $targetMailbox -User $delegate -automapping $false -AccessRights FullAccess -InheritanceType All -ErrorAction Stop
                                }

                                $msg = "      SUCCESS: FullAccess permission applied."
                                Write-Host -ForegroundColor Green $msg
                                Log-Write -Message $msg
                                $FullAccessPermissionsCount += 1   
                            }
                            catch {
                                $msg = "      ERROR: Failed to apply FullAccess permission."
                                Write-Host -ForegroundColor Red  $msg
                                Write-Host -ForegroundColor Red $_.Exception.Message
                                Log-Write -Message $msg
                                Log-Write -Message $_.Exception.Message
                            }                    
                        }
                    }
                    else {
                        $msg = "      WARNING: The verification status for FullAccess permission is still in pending in G Suite." 
                        Write-Host -ForegroundColor Yellow $msg 
                        Log-Write -Message $msg
                        $msg = "      WARNING: FullAccess permission won't be re-applied."   
                        Write-Host -ForegroundColor Yellow $msg 
                        Log-Write -Message $msg         
                    }
                }
                else {
                    $msg = "      ERROR: Delegate '$delegate' doest not exist in Office 365. FullAccess permission skipped."
                    Write-Host -ForegroundColor Red  $msg
                    Log-Write -Message $msg
                }
        
            }
            else{
                $msg =  "      ERROR: Target mailbox '$targetMailbox' doest not exist in Office 365."
                Write-Host -ForegroundColor Red  $msg
                Log-Write -Message $msg
            }  
        }

        if($FullAccessPermissionsCount -ge 2) {
            $msg = "SUCCESS: $FullAccessPermissionsCount FullAccess permissions out of $totalFullAccessPermissionsExport have been applied to Office 365 mailboxes."
            Write-Host -ForegroundColor Green $msg
            Log-Write -Message $msg
        }elseif ($FullAccessPermissionsCount -eq 1) {
            $msg = "SUCCESS: 1 FullAccess permission out of $totalFullAccessPermissionsExport has been applied to Office 365 mailboxes."
            Write-Host -ForegroundColor Green $msg
            Log-Write -Message $msg
        }
    }
}

#######################################################################################################################
#                                        IMPORT CALENDAR PERMISSIONS INTO O365
#######################################################################################################################
Function Process-CalendarPermissions {

    Write-Host "INFO: Exporting Calendar permissions to CSV file." 

    #Export calendarPermissionsArray to CSV file
    try {
        if($onlyPermissionsReport) {
            $script:calendarPermissionsArray | Export-Csv -Path $workingDir\GSuiteCalendarACLReport.csv -NoTypeInformation -force
            Write-Host -ForegroundColor Green "SUCCESS: CSV file '$workingDir\GSuiteCalendarACLReport.csv' processed, exported and open."
        }
        else {
            $script:calendarPermissionsArray | Export-Csv -Path $workingDir\GSuiteCalendarACL.csv -NoTypeInformation -force
            Write-Host -ForegroundColor Green "SUCCESS: CSV file '$workingDir\GSuiteCalendarACL.csv' processed, exported and open."
        }   

        if ($script:sameEmailAddresses) {
            Write-Host -ForegroundColor Yellow "         ACTION:  Please review the opened CSV file and once you finish, save it."
        }
        elseif(!$script:sameEmailAddresses -and $script:sameUserName -and $script:destinationDomain -ne "" -and $userDomain -eq $calendarDomain -and $onlyPermissionsReport -eq $false) {
            Write-Host -ForegroundColor Yellow "         WARNING: The 'user' column of the opened CSV file has been updated with the new domain."
            Write-Host -ForegroundColor Yellow "         WARNING: The 'delegateAddress' column of the opened CSV file has been updated with the new domain."
            Write-Host -ForegroundColor Yellow "         ACTION:  Please review the opened CSV file and once you finish, save it."
        }
        elseif(!$script:sameEmailAddresses -and $script:sameUserName -and $script:destinationDomain -ne ""  -and $userDomain -ne $calendarDomain -and $onlyPermissionsReport -eq $false) {
            Write-Host -ForegroundColor Yellow "         WARNING: The 'user' column of the opened CSV file has been updated with the new domain."
            Write-Host -ForegroundColor Yellow "         ACTION:  Populate the 'destinationDelegateAddress' column of the opened CSV file with the destination Calendar email."
            Write-Host -ForegroundColor Yellow "         ACTION:  Please review the opened CSV file and once you finish, save it."
        }
        elseif(!$script:sameEmailAddresses -and !$script:sameUserName -and $script:destinationDomain -ne "" -and $onlyPermissionsReport -eq $false) {
            Write-Host -ForegroundColor Yellow "         ACTION: Populate the 'destinationUser' column of the opened CSV file with the destination user email."
            Write-Host -ForegroundColor Yellow "         ACTION: Populate the 'destinationDelegateAddress' column of the opened CSV file with the destination Calendar email."
            Write-Host -ForegroundColor Yellow "         ACTION: Once you finish editing the CSV file, save it."
        }
        elseif($onlyPermissionsReport -eq $false) {
        }
    }
    catch {
        if($onlyPermissionsReport) {
            $msg = "ERROR: Failed to import the CSV file '$workingDir\GSuiteCalendarACLReport.csv'."
        }
        else {
            $msg = "ERROR: Failed to import the CSV file '$workingDir\GSuiteCalendarACL.csv'."
        }         
        Write-Host -ForegroundColor Red  $msg
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $msg
        Log-Write -Message $_.Exception.Message
        Exit
    }

    #Open the CSV file for editing
    if($onlyPermissionsReport) {
        Start-Process -FilePath $workingDir\GSuiteCalendarACLReport.csv
    }
    else {
        Start-Process -FilePath $workingDir\GSuiteCalendarACL.csv
    }

    #If the script must generate GSuite permissions report and also migrate them to O365
    if(!$onlyPermissionsReport) {
        Write-Host "ACTION: If you have reviewed, edited and saved the CSV file then press any key to continue." 
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');

        #Re-import the edited CSV file
        Try{
            $calendarPermissions = @(Import-CSV "$workingDir\GSuiteCalendarACL.csv" | where-Object { $_.PSObject.Properties.Value -ne ""}) 
            #Write-Host -ForegroundColor Green "SUCCESS: CSV file '$workingDir\GSuiteCalendarACL.csv' re-imported."
        }
        Catch [Exception] {
            $msg = "ERROR: Failed to import the CSV file '$workingDir\GSuiteCalendarACL.csv'. Please save and close the CSV file."
            Write-Host -ForegroundColor Red  $msg
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $msg
            Log-Write -Message $_.Exception.Message
            Exit
        }

        Write-Host
        $confirm = (Read-Host -prompt "Do you want to send an email with the published calendar URL to external users?  [Y]es or [N]o")
        if($confirm.ToLower() -eq "y") {
            $confirmExternalPermissions = $true
        }

        $totalCalendarPermissionsExport = $calendarPermissions.count
        $calendarPermissionsCount = 0
        $publishedCalendarUrlCount = 0
        $currentCalendarPermission = 0

        Write-Host 
        Write-Host "INFO: Importing Calendar permissions into Office 365."  
        Foreach($calendarPermission in $calendarPermissions){

            $currentCalendarPermission += 1 
            if($calendarPermission.destinationCalendar) {
                $calendar = $calendarPermission.destinationCalendar
            }
            else {
                $calendar = $calendarPermission.calendar
            }

            if($calendarPermission.destinationScope) {
                $scope = $calendarPermission.destinationScope
            }
            else {
                $scope = $calendarPermission.scope
            }

            $scopeType = $calendarPermission.scopeType
            $o365Role = $calendarPermission.o365Role

            $msg = "INFO: Processing Calendar permission $currentCalendarPermission/$totalCalendarPermissionsExport : Calendar $calendar Delegate $scope Role $o365Role."
            Write-Host -ForegroundColor Gray $msg
            Log-Write -Message $msg

            #Verify if target mailbox exists
            $recipient = Get-Recipient -identity $calendar -ErrorAction SilentlyContinue
    
            if ($recipient -ne $null) {
            
                #Verify if scope exists
                $isInternalScope = check-O365User -user $scope
                #Verify if mailboxFolderPermission exists
                $folderPermission = Get-MailboxFolderPermission -Identity $calendar":\calendar" -user $scope  -ErrorAction SilentlyContinue
                $scopeDomain = $scope.split("@")[1]

                if($folderPermission) {
                        $msg = "      WARNING: Calendar permission already exists in Office 365."
                        Write-Host -ForegroundColor Yellow $msg
                        Log-Write -Message $msg
                }
                elseif(!$folderPermission -and $isInternalScope) {
            
                        $result = Add-MailboxFolderPermission -Identity $calendar":\calendar" -user $scope -AccessRights $o365Role -ErrorAction SilentlyContinue

                        if ($result) {
                            $msg = "      SUCCESS: Calendar permission applied."
                            Write-Host -ForegroundColor Green $msg
                            Log-Write -Message $msg
                            $calendarPermissionsCount += 1
                        }
                        else {
                            $msg = "      ERROR: Calendar permission not applied."
                            Write-Host -ForegroundColor Red $msg
                            Log-Write -Message $msg
                            $calendarPermissionsCount += 1
                        }
                }
                elseif(!$folderPermission -and !$isInternalScope -and ($scopeDomain -eq $script:destinationDomain)) {
                    $msg =  "      ERROR: Scope '$scope' doest not exist in Office 365."
                    Write-Host -ForegroundColor Red  $msg
                    Log-Write -Message $msg
                }
                elseif(!$folderPermission -and !$isInternalScope -and ($scopeDomain -ne $script:destinationDomain)) {
                    $msg = "      WARNING: User '$scope' does not exist in Office 365, is an external user." 
                    Write-Host -ForegroundColor Yellow $msg
                    Log-Write -Message $msg

                    if($confirmExternalPermissions) {
                        
                        $publishedCalendar = Get-MailboxCalendarFolder -Identity "$($recipient.Identity):\calendar" 
                        if (!$publishedCalendar.publishEnabled) {
                            $result = Set-MailboxCalendarFolder -Identity "$($recipient.Identity):\calendar" -DetailLevel AvailabilityOnly -PublishEnabled $true -ErrorAction SilentlyContinue

                            if ($result) {
                                $msg = "      SUCCESS: Calendar Sharing URL published."
                                Write-Host -ForegroundColor Green $msg
                                Log-Write -Message $msg

                                $publishedCalendar = Get-MailboxCalendarFolder -Identity "$($recipient.Identity):\calendar" -ErrorAction SilentlyContinue
                            }
                            else {
                                $msg = "      ERROR:  Calendar Sharing URL not published."
                                Write-Host -ForegroundColor Red $msg
                                Log-Write -Message $msg
                                Continue
                            }
                        }                        

                        #published calendar HTML URL
                        $htmlPublishedCalendarUrl=$publishedCalendar.PublishedCalendarUrl
                        #published calendar ICS URL
                        $icsPublishedCalendarURL = $publishedCalendar.PublishedICalUrl 
                        $smtpServer = "smtp.office365.com"
                        $smtpCreds = $destinationO365Creds
                        $emailTo = $scope
                        $emailFrom = $smtpCreds.Username
                        $FirstName = $recipient.FirstName
                        $LastName = $recipient.LastName
                        $subject = "You're invited to share this calendar"
                        $body = ""
                        $body += "<h2>I'd like to share my calendar with you </h2><br>"
                        $body += "$FirstName $LastName ($calendar) would like to share an Outlook calendar with you. <br><br>"
                        $body += "You'll be able to see the availability information of events on <a href="+$htmlPublishedCalendarUrl+">this calendar</a>. <br><br>"
                        $body += "To import the calendar to your Outlook calendar, this is the <a href="+$icsPublishedCalendarURL+">ICS file</a>."

                        try {
                            $result = Send-MailMessage -To $emailTo -From $emailFrom -Subject $subject -Body $body -BodyAsHtml -SmtpServer $smtpServer -Port 587 -Credential $smtpCreds -UseSsl -ErrorAction SilentlyContinue
                        
                            if ($error[0].ToString() -match "Spam abuse detected from IP range.") { 
                                $msg = "      ERROR: Failed to send email to user '$emailTo'. Access denied, spam abuse detected. The sending account has been banned. "
                                Write-Host -ForegroundColor Red  $msg
                                Log-Write -Message $msg
                            }
                            else {
                                $msg = "      SUCCESS: Email with $FirstName's calendar URL sent to external user '$emailTo'"
                                Write-Host -ForegroundColor Green $msg
                                Log-Write -Message $msg 
                                $publishedCalendarUrlCount += 1   
                           }
  
                        }
                        catch {
                            $msg = "      ERROR: Failed to send email to user '$emailTo'."
                            Write-Host -ForegroundColor Red  $msg
                            Write-Host -ForegroundColor Red $_.Exception.Message
                            Log-Write -Message $msg
                            Log-Write -Message $_.Exception.Message
                        }

                    }
                }
        
            }
            else{
                $msg =  "      ERROR: Target mailbox '$calendar' doest not exist in Office 365."
                Write-Host -ForegroundColor Red  $msg
                Log-Write -Message $msg
            }    
        }

        if($calendarPermissionsCount -ge 2) {
            $msg = "SUCCESS: $calendarPermissionsCount Calendar permissions out of $totalCalendarPermissionsExport have been applied to Office 365 mailboxes."
            Write-Host -ForegroundColor Green $msg
            Log-Write -Message $msg
        }elseif ($calendarPermissionsCount -eq 1) {
            $msg = "SUCCESS: 1 Calendar permission out of $totalCalendarPermissionsExport has been applied to Office 365 mailboxes."
            Write-Host -ForegroundColor Green $msg
            Log-Write -Message $msg
        }

        if($publishedCalendarUrlCount -ge 2) {
            $msg = "SUCCESS: $publishedCalendarUrlCount published Calendar URLs have been sent to external users."
            Write-Host -ForegroundColor Green $msg
            Log-Write -Message $msg
        }elseif ($publishedCalendarUrlCount -eq 1) {
            $msg = "SUCCESS: 1 published Calendar URL has been sent to an external user."
            Write-Host -ForegroundColor Green $msg
            Log-Write -Message $msg
        }
    }
}


#######################################################################################################################
#                                               MAIN PROGRAM
#######################################################################################################################

## Initiate Parameters

#Working Directory
$workingDir = "C:\scripts"

#Logs directory
$logDirName = "LOGS"
$logDir = "$workingDir\$logDirName"

#Log file
$logFileName = "$(Get-Date -Format yyyyMMdd)_Migrate-PermissionsGSuiteO365.log"
$logFile = "$logDir\$logFileName"

Create-Working-Directory -workingDir $workingDir -logDir $logDir

Write-Host 
Write-Host -ForegroundColor Yellow "WARNING: Minimal output will appear on the screen." 
Write-Host -ForegroundColor Yellow "         Please look at the log file '$($logFile)'."
Write-Host -ForegroundColor Yellow "         All CSV files will be in folder '$($workingDir)'."
Write-Host 
Start-Sleep -Seconds 1

$msg = "++++++++++++++++++++++++++++++++++++++++ SCRIPT STARTED ++++++++++++++++++++++++++++++++++++++++"
Log-Write -Message $msg

#To only have source email addresses when only generating reports
$script:sameEmailAddresses = $true
$createUserBatches = $false

#Main menu
do {
    $confirm = (Read-Host -prompt "What do you want to do? `
    1. Generate GSuite permissions [r]eports `
    2. Generate user [b]atches based on delegate permissions `
    3. [M]igrate mailbox and folder permissions to Office 365`
    4. [E]xit

[R]eports, [B]atches, [M]igration or [E]xit")
} while(($confirm.ToLower() -ne "r") -and ($confirm.ToLower() -ne "m") -and ($confirm.ToLower() -ne "b") -and ($confirm.ToLower() -ne "e"))

if($confirm.ToLower() -eq "r") {
    $onlyPermissionsReport=$true
    Write-Host
    Write-Host -ForegroundColor Gray "INFO: This script is going to only export mailbox and/or folder permissions from source GSuite to CSV files." 
    Write-Host    
    $destinationO365Session = Connect-ExchangeOnlineDestination
    Write-Host
}
elseif ($confirm.ToLower() -eq "b") {
    $onlyPermissionsReport=$false
    Write-Host
    Write-Host -ForegroundColor Gray "INFO: This script is going to export all FullAccess mailbox permissions from source GSuite"
    Write-Host -ForegroundColor Gray "      and generate user batches based on these exported FullAccess permissions." 
    Write-Host    
    $destinationO365Session = Connect-ExchangeOnlineDestination
    Write-Host
    $createUserBatches = $true
}
elseif ($confirm.ToLower() -eq "m") {
    $onlyPermissionsReport=$false
    Write-Host
    Write-Host -ForegroundColor Gray "INFO: This script is going to export mailbox and/or folder permission from source GSuite to CSV files"
    Write-Host -ForegroundColor Gray "      and import them into destination Office 365." 
    Write-Host    
    $destinationO365Session = Connect-ExchangeOnlineDestination
    Write-Host
    query-EmailAddressMapping
    Write-Host
}
elseif ($confirm.ToLower() -eq "e") {
    Exit
}

$skipNonExistingUser=$false
do {
    $confirm = (Read-Host -prompt "Do you want to skip the users that do not exist in destination Office 365?  [Y]es or [N]o")
} while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))

if($confirm.ToLower() -eq "y") {
    $skipNonExistingUser=$true
}

$readCSVFile=$false
do {
    $confirm = (Read-Host -prompt "Do you want to import a CSV file with the users you want to process?  [Y]es or [N]o")
} while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))

if($confirm.ToLower() -eq "y") {
    $readCSVFile=$true
}

if($readCSVFile) {
	Write-Host -ForegroundColor yellow "ACTION: Select the CSV file to import file (Press cancel to create one)"
    Get-FileName $workingDir
}

if($createUserBatches -eq $false) {
    $processSendAs = $false
    $processFullAccess = $false
    $processCalendars = $false
    $processOnlyCalendars = $false

    Write-Host
    do {
        if ($onlyPermissionsReport) {
            $confirm = (Read-Host -prompt "Do you want to generate source GSuite SendAs permissions report?  [Y]es or [N]o")    
        }
        else {
            $confirm = (Read-Host -prompt "Do you want to migrate SendAs permissions from source GSuite to destination Office 365?  [Y]es or [N]o")
        }
    } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))

    if($confirm.ToLower() -eq "y") {
        $processSendAs = $true
    }

    do {
        if ($onlyPermissionsReport) {
            $confirm = (Read-Host -prompt "Do you want to generate source GSuite FullAccess permissions report?  [Y]es or [N]o")    
        }
        else{
            $confirm = (Read-Host -prompt "Do you want to migrate FullAccess permissions from source GSuite to destination Office 365?  [Y]es or [N]o")
        }
    } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))

    if($confirm.ToLower() -eq "y") {
        $processFullAccess = $true
    }

    do {
        if ($onlyPermissionsReport) {
            $confirm = (Read-Host -prompt "Do you want to generate source GSuite Calendar permissions report?  [Y]es or [N]o")   
            $processOnlyCalendars = $true
        }
        else{
            $confirm = (Read-Host -prompt "Do you want to migrate Calendar permissions from source GSuite to destination Office 365?  [Y]es or [N]o")
            $processOnlyCalendars = $true
        }
    } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))

    if($confirm.ToLower() -eq "y") {
        $processCalendars = $true
    } 
    
}
else {
    #For user batch creation, FullAccess permissions are taken into consideration.
    $processSendAs = $false
    $processFullAccess = $true
    $processCalendars = $false
}

$result = Export-GsuitePermissions -skipNonExistingUser $skipNonExistingUser -processSendAs $processSendAs -processFullAccess  $processFullAccess -processCalendars $processCalendars -userBatches $createUserBatches -readCSVfile $readCSVfile 

if($processSendAs -and $createUserBatches -eq $false) {
    Write-Host
    Process-SendAsPermissions 
    Start-Sleep -s 5
}
if($processFullAccess -and $createUserBatches -eq $false) {
    Write-Host
    Process-FullAccessPermissions
    Start-Sleep -s 5
}
if($processCalendars -and $createUserBatches -eq $false) {
    Write-Host
    Process-CalendarPermissions
    Start-Sleep -s 5
}


Write-Host
try {
    Write-Host "INFO: Opening directory $workingDir where you will find all the generated CSV files."
    Invoke-Item $workingDir
    Write-Host
}
catch{
    $msg = "ERROR: Failed to open directory '$workingDir'. Script will abort."
    Write-Host -ForegroundColor Red $msg
    Exit
}

$msg = "++++++++++++++++++++++++++++++++++++++++ SCRIPT FINISHED ++++++++++++++++++++++++++++++++++++++++`n"
Log-Write -Message $msg

Remove-PSSession $destinationO365Session

##END SCRIPT

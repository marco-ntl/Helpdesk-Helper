# GLOBAL CONST.
$EXPORT_PATH = "C:\Users\adm-DaboDa\Desktop\EXPORT\"
$LOGS_PATH = "E:\PublicLogs\"
$NOW = (Get-Date).toString("dd.MM.yy")

function Compare-PasswordlastSet {

    # VAR & CONST
    $FIRST_CONTACT_WEAK = Get-Date -Date "07.07.2023 00:00:00"
    $funcName = $MyInvocation.MyCommand.Name
    $COUNT = 0

    # Import data form csv file
    $fileName = Get-FilePathFromDialog
    $csvData = Import-Csv -Path $fileName -Delimiter ";" -Encoding "utf8"

    foreach ($item in $csvData) {

        $userName = $item.SamAccountName

        # Retrieve the user object from Active Directory
        $user = Get-ADUser -Identity $userName -Properties pwdLastSet

        # Extract the pwdLastSet value
        $pwdLastSet = $user.pwdLastSet
        $pwdLastSetDateTime = [DateTime]::FromFileTime($pwdLastSet)

        # Compare the two DateTime values
        if ($pwdLastSetDateTime -gt $FIRST_CONTACT_WEAK) {
            Write-Success "password changed at: $pwdLastSetDateTime for user: $username"
            Write-CustomLogs -From $funcName "YES - password changed after comm on 07.07" -Type INFO
            $COUNT ++
        } 
        else {
            Write-Err "No change password for: $userName"
            Write-CustomLogs -From $funcName "NO - no change password" -Type INFO
        }  
    }
    Write-CustomLogs -From $funcName "Result: $COUNT - passwords changed" -Type INFO
}
function Disable-UsersSecurityActions {

    # Import data form csv file
    $fileName = Get-FilePathFromDialog
    $csvData = Import-Csv -Path $fileName -Delimiter ";" -Encoding "utf8"

    # Disable user
    foreach ($item in $csvData) {

        $userName = $item.SamAccountName

        try {
            Set-ADUser -identity $userName  -Enable $false 
            Write-Success "$NOW - Disabled for user: $userName"
        }

        catch {
            Write-Err $_.Exception.Message
        }
    }  
}
function Export-RTSExternalUsersAndLicenceO365 {

    # VAR & CONST
    $FILENAME = $MyInvocation.MyCommand.Name + ".csv" 
    $COMPLETE_PATH = $EXPORT_PATH + $NOW + "-" + $FILENAME
    $ErrorActionPreference = "silentlycontinue"                     

    $externalUserList = (Get-ADUser -filter * -SearchBase "OU=2_current,OU=External,OU=Users,OU=RTS,OU=Units,DC=media,DC=int" -Properties samAccountName | Select-Object samAccountName)
    Write-Info "External Users list OK"

    foreach ($item in $externalUserList) {

        $samAccountName = $item.sAMAccountName
        $licenceType = (Get-ADUser -Identity $samAccountName -Properties MemberOf | Select-Object MemberOf).MemberOf | Where-Object {$_ -like "*RTS-G-L-M365*"}
        $licenceType = $licenceType.Split(",")[0].trim("CN=")

        Get-ADUser -identity $samAccountName -Properties  mail, manager ,accountExpirationDate ,sn ,givenName ,employeeID ,title ,division ,employeeType ,department ,lastLogonDate ,mobile |
        Select-Object mail, manager ,accountExpirationDate ,sn ,givenName ,employeeID ,title ,division ,employeeType ,department ,lastLogonDate ,mobile, @{name="O365 Licence"; expression={$licenceType}} |
        Export-Csv -Path $COMPLETE_PATH -NoTypeInformation -Delimiter ";" -Append -Encoding "utf8"
    }

    Write-Info "External Users exported..."
}
Function Export-RTSConferenceRooms {

    # VAR & CONST
    $FILENAME = $MyInvocation.MyCommand.Name + ".csv"
    $COMPLETE_PATH = $EXPORT_PATH + $NOW + '-' + $FILENAME
    
    Import-O365Session -silent:$true

    $Rooms = Get-Mailbox -ResultSize unlimited -Filter "(RecipientTypeDetails -eq 'RoomMailbox') -and (Alias -like 'RTS-*') -and (name -notlike '*test*')" | Select-Object Alias | Sort-Object Alias
    
    foreach ($item in $Rooms) {

        $Alias = $item.Alias
        $Room = Get-CalendarProcessing -identity $Alias | Select-Object -Property @{Name = 'Alias'; Expression = { $Alias } }, Identity, ResourceDelegates, BookInPolicy, AllBookInPolicy, AllRequestInPolicy, AutomateProcessing, BookingWindowInDays, MaximumDurationInMinutes, ConflictPercentageAllowed, MaximumConflictInstances, RemovePrivateProperty, DeleteSubject, AddOrganizerToSubject, AddAdditionalResponse, AdditionalResponse
        $Room | Export-Csv -Path $COMPLETE_PATH -NoTypeInformation -Delimiter ";" -Append -Encoding "utf8"    
    }
}
function Export-RTSComputers {

    # VAR & CONST
    $FILENAME = $MyInvocation.MyCommand.Name + ".csv"
    $COMPLETE_PATH = $EXPORT_PATH + $NOW + '-' + $FILENAME
 
    Get-ADComputer -Filter * -Properties cn, Description, distinguishedName, Enabled, lastLogonDate, operatingSystem | Select-Object cn, Description, distinguishedName, Enabled, lastLogonDate, operatingSystem  `
    | Export-Csv -Path $COMPLETE_PATH -NoTypeInformation -Delimiter ";" -Force -Encoding "utf8"

    Write-Info "Finished... CSV File located in: $EXPORT_PATH"
}
function Get-PasswordNeverExpires {

    # VAR & CONST
    $FILENAME = $MyInvocation.MyCommand.Name + ".csv"
    $COMPLETE_PATH = $EXPORT_PATH + $NOW + '-' + $FILENAME

    get-aduser -filter * -Properties samAccountName, distinguishedName, EmailAddress, PasswordNeverExpires |
    Select-Object samAccountName, distinguishedName, EmailAddress, PasswordNeverExpires | Where-Object { $_.distinguishedName -like "*OU=2_current,OU=Internal,OU=Users,OU=RTS,OU=Units,DC=media,DC=int*" -or $_.distinguishedName -like "*OU=2_current,OU=External,OU=Users,OU=RTS,OU=Units,DC=media,DC=int*" } |
    Export-Csv -Path $COMPLETE_PATH -NoTypeInformation -Delimiter ";" -Append -Encoding "utf8"
}
<#
    .SYNOPSIS
    This function import data from csv file to do Security Taskforce Action

    .DESCRIPTION 
    This function import data from  a .csv file then put them in a loop and iterate on each item to:
    - Ping the item with echo protocol -> if they return false continue, else skip tasks
    - Move them from an OU to another
    - Rename the item in field "Description"
    - Disable the item in AD
#>
function Invoke-ComputersSecurityActions {

    
        # VAR & CONST
        $DEST_OU = 'OU=Z_Nettoyage,OU=Computers,OU=RTS,OU=Units,DC=media,DC=int'
        $NEW_DESCRIPTION = "Security Taskforce Action - DDA - $NOW"
        $funcName = $MyInvocation.MyCommand.Name
        $computerToChange = [System.Collections.ArrayList]::new()
        $count = 0
    
        # Start import data form csv file
        $fileName = Get-FilePathFromDialog
        $csvData = Import-Csv -Path $fileName -Delimiter ";" -Encoding "utf8"
        Write-Info "data imported, starting process..."
    
        # First Loop to extract ping test
        foreach ($item in $csvData) {
    
            $deviceName = $item.CN
    
            # if ping is True
            if (Test-Connection -ComputerName $deviceName -Quiet -Count 1) {
    
                Write-Err "$deviceName :: Respond to ping - No Actions will be taken"
                Write-CustomLogs -From $funcName "$deviceName :: Respond to ping - No Actions will be taken" -Type WARNING
            }
            # else ping is false
            else {
    
                Write-Success "$deviceName :: NO Respond to ping - Quarantine can continue"
                Write-CustomLogs -From $funcName "$deviceName :: NO Respond to ping - Quarantine can continue" -Type SUCCESS
    
                $computerToChange.Add("$deviceName")
            }
        }
    
        Write-Info "Display device to take security actions:: `n$computerToChange"
    
        Write-Success "Data exported in C:\Users\Public\PublicLogs Folder ...`ntake security actions to change password ?"
        $confirm = Read-Host "[Y]es or [N]o"
    
        if ($confirm -eq "y" -or $confirm -eq "Y") {
            
            # 2nd Loop to manage Security Actions
            foreach ($item in $computerToChange) {
    
                $deviceNameToGo = $computerToChange[$count]
                Write-Info "Actions executed for device: $deviceNameToGo"
                $count += 1
    
                try {
                    # Step 2: Disable device 
                    Disable-ADAccount -Identity $deviceNameToGo
    
                    # Step 3: Set description 
                    Set-ADObject -Identity $deviceNameToGo -Description $NEW_DESCRIPTION
    
                    # Step 4: Move device to OU: media.int/Units/RTS/Computers/Z_Nettoyage
                    Move-ADObject -Identity $deviceNameToGo -TargetPath $DEST_OU  
    
                    Write-Success "Quarantine actions completed for device: $deviceNameToGo"
                    Write-CustomLogs -From $funcName "Quarantine actions completed for device: $deviceNameToGo" -Type SUCCESS
                }
    
                catch {
                    Write-Err "Action cannot be completed for device: $deviceNameToGo"
                    Write-Err $_
                    Write-CustomLogs -From $funcName "$_ / Error for: $deviceNameToGo" -Type ERROR
                    continue
                }
            }
        }
        if (Ask-Confirmation "Recommencer ?") {
            Invoke-ComputersSecurityActions
        }
}   
function Invoke-UsersSecurityActions {

    # VAR & CONST
    $FILENAME = $MyInvocation.MyCommand.Name + ".csv"
    $COMPLETE_PATH = $EXPORT_PATH + $NOW + '-' + $FILENAME

    # Import data form csv file
    $fileName = Get-FilePathFromDialog
    $csvData = Import-Csv -Path $fileName -Delimiter ";" -Encoding "utf8"
    Write-Host "data imported, starting process..."
    foreach ($item in $csvData) {

        $userName = $item.SamAccountName

        # Get data and export to csv
        Get-ADUser -identity $userName -Properties samAccountName, distinguishedName, Enabled, PasswordLastSet, EmailAddress, extensionAttribute10, AccountExpirationDate, lastLogonDate |
        Select-Object samAccountName, distinguishedName, Enabled, PasswordLastSet, EmailAddress, extensionAttribute10, AccountExpirationDate, lastLogonDate |
        Export-Csv -Path $COMPLETE_PATH -NoTypeInformation -Delimiter ";" -Append -Encoding "utf8"
    }

    Write-Success "Data exported in D:\ Public Folder ...`ntake security actions to change password ?"
    $confirm = Read-Host "[Y]es or [N]o"

    if (($confirm -eq "y") -or ($confirm -eq "Y")) {

        # Force change password at next logon
        foreach ($item in $csvData) {

            $userName = $item.SamAccountName

            try {
                Set-ADUser -identity $userName -PasswordNeverExpires:$false 
                Set-ADUser -identity $userName -ChangePasswordAtLogon:$true 
                Write-Success "Force change password at next logon for user: $userName"
                Write-CustomLogs -From $funcName "Force change password at next logon for user: $userName" -Type INFO   
            }

            catch {
                Write-Err "$_.Exception.Message"
                Write-CustomLogs -From $funcName "$_.Exception.Message" -Type ERROR             
            }
        }  
    }

    else {
        Exit
    } 
}  
function Search-MailboxAccessAudit ([string]$username, [string]$type) {

    # TODO: Add filter for search on RTS mailboxes only

    Import-O365Session -silent:$true

    $type = Assert-StrArg $type '' "Type de Mailbox: [SHARED] / [USER]"
    $type = $type + "Mailbox" 

    $MBXs = Get-Mailbox -RecipientTypeDetails $type -ResultSize Unlimited
    Write-Host "$type" + "Mailboxes loaded"

    Foreach ($MBX in $MBXs) {

        Write-Host "$MBX :: Testing..."
        Get-MailboxPermission -identity $MBX -user $username
    }
}
function Set-AccountExpires {

    # Import data form csv file
    $NOW = (Get-Date).toString("dd.MM")
    $fileName = Get-FilePathFromDialog
    $csvData = Import-Csv -Path $fileName -Delimiter ";" -Encoding "utf8"

    # Disable user
    foreach ($item in $csvData) {

        $userName = $item.distinguishedName

        try {
            set-aduser -identity $userName -Property accountExpires 0
            Write-Success "$NOW - accountExpires set to 0 for user: $userName"
        }

        catch {
            Write-Err "$_.Exception.Message"
        }
    }  
}
function Write-CustomLogs {

    Param (
        [parameter(Mandatory = $true)]
        [string]$From,
        [parameter(Mandatory = $true)]
        [string]$LogString,
        [parameter(Mandatory = $true)]
        [ValidateSet("WARNING", "ERROR", "INFO", "SUCCESS")]
        [string]$Type
    )

    $NOW = (Get-Date).toString("dd.MM.yy")
    $LogfileName = $From 
    $LogfilePath = $LOGS_PATH + $NOW + '-' + $LogfileName + '.log'
    
    $TimeStamp = (Get-Date).toString("dd.MM.yyyy HH:mm:ss")
    $LogMessage = $TimeStamp + " - " + $type + " - " + $LogString
    Add-content $LogfilePath -value $LogMessage
}

# Other Stuffs
<# 
    *** RECEVEID FROM GD *** 
    ce truc ne fonctionne pas c'est de la daube.... => pas les droits pour query les DomainsControllers ***

    .SYNOPSIS 
    This function will locate the computer that processed a failed user logon attempt which caused the user account to become locked out. 
 
    .DESCRIPTION 
    This function will locate the computer that processed a failed user logon attempt which caused the user account to become locked out.  
    The locked out location is found by querying the PDC Emulator for locked out events (4740).   
    The function will display the BadPasswordTime attribute on all of the domain controllers to add in further troubleshooting. 
 
    .EXAMPLE 
    PS C:\>Get-LockedOutLocation -Identity Joe.Davis 
    This example will find the locked out location for Joe Davis. 

    .NOTES
    This function is only compatible with an environment where the domain controller with the PDCe role to be running Windows Server 2008 SP2 and up.   
    The script is also dependent the ActiveDirectory PowerShell module, which requires the AD Web services to be running on at least one domain controller. 
    Author:Jason Walker 
    Last Modified: 3/20/2013 
#> 
Function Get-LockedOutLocation { 

    [CmdletBinding()] 
 
    Param( 
        [Parameter(Mandatory = $True)] 
        [String]$Identity       
    ) 
 
    Begin {  
        $DCCounter = 0  
        #$LockedOutStats = @()    
        Clear-Host         
        
        Try { 
            Import-Module ActiveDirectory -ErrorAction Stop 
        } 
        Catch { 
            Write-Warning $_ 
            Break 
        } 
    }#end begin 
    Process { 
         
        #Get all domain controllers in domain 
        $DomainControllers = Get-ADDomainController -Filter * 
        $PDCEmulator = ($DomainControllers | Where-Object { $_.OperationMasterRoles -contains "PDCEmulator" }) 
        
        Foreach ($DC in $DomainControllers) { 
            $DCCounter++ 
            Write-Progress -Activity "Contacting DCs for lockout info" -Status "Querying $($DC.Hostname)" -PercentComplete (($DCCounter / $DomainControllers.Count) * 100) 
            Try { 
                $UserInfo = Get-ADUser -Identity $Identity -Server $DC.Hostname -Properties LastBadPasswordAttempt -ErrorAction Stop 
            } 
            Catch { 
                Write-Warning $_ 
                Continue 
            } 
            If ($UserInfo.LastBadPasswordAttempt) {     
                $SID = $UserInfo.SID.Value          
            }#end if 
        } 
        #Get User Info 
        Try {   
            Write-Verbose "Querying event log on $($PDCEmulator.HostName)" 
            Write-Progress -Activity "Querying event log on $($PDCEmulator.HostName)"
            $LockedOutEvents = Get-WinEvent -ComputerName $PDCEmulator.HostName -FilterHashtable @{LogName = 'Security'; Id = 4740 } -ErrorAction Stop | Sort-Object -Property TimeCreated -Descending 
        } 
        Catch {           
            Write-Warning $_ 
            Continue 
        }#end catch      
                                  
        $lockouts = Foreach ($Event in $LockedOutEvents) {             
            $Eventcounter++
            Write-Progress -Activity "Collecting Lockout Events from Security Log" -PercentComplete (($Eventcounter / $LockedOutEvents.Count) * 100)
            If ($Event | Where-Object { $_.Properties[2].value -match $SID }) {  
               
                $Event | Select-Object -Property @( 
                
                    @{Label = 'LockedOutLocation'; Expression = { $_.Properties[1].Value } } 
                ) 
                                                 
            }#end ifevent 
             
        }#end foreach lockedout event 
        
        $lockouts | Sort-Object -Property LockedOutLocation -Unique            
                        
    }#end process 
    
}



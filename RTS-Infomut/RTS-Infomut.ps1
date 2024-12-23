using module RTS-Components
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.OpenFileDialog")
$DAYS_BEFORE_NETTOYAGE = 3
$DAYS_BEFORE_LEAVERS = 7

$RTS_OU = 'OU=RTS,OU=Units,DC=media,DC=int'
$RTS_USERS_OU = "OU=Users,$RTS_OU"
$INTERNAL_OU = "OU=2_Current, OU=Internal,$RTS_USERS_OU"
$EXTERNAL_OU = "OU=2_Current, OU=External,$RTS_USERS_OU"

$NETTOYAGE_OU = "OU=1 Infomut_Expires,OU=13-Nettoyage,$($RTS_USERS_OU)"
$LEAVERS_OU = "OU=Leavers,$($RTS_USERS_OU)"

$EXCEPTION_KEYWORD = "NoLeavers"

class Leaver{
    $ADUser
    [DateTime]$expDate
    [string]$ticketNumber

    Leaver($user, [DateTime]$expDate, [string]$ticketNumber){
        $this.ADUser = $user
        $this.expDate = $expDate
        $this.ticketNumber = $ticketNumber
    }
}
Function Set-InfomutException([string]$user, $NoLeavering = $null, [string]$ticket){
    $user = Assert-StrArg $user $_ "Nom d'utilisateur"
    if($null -eq $NoLeavering){
        $NoLeavering = Ask-Confirmation "L'utilisateur doit-il faire parties des exceptions Infomut ?"
    }
    $ticket = Assert-StrArg $ticket '' "Numéro de ticket"

    $newDesc = "[$(Get-Date -format dd.MM)] $ticket"
    if($NoLeavering){
        Set-ADUser $user -Replace @{ExtensionAttribute9=$EXCEPTION_KEYWORD; Description="$newDesc - Ajouté aux exception Infomut"} -Clear ExtensionAttribute11, ExtensionAttribute12
        Write-Success "User $($user.SamAccountName) ne sera plus affecté par le processus infomut"
    }else{
        Set-ADUser $user -Clear ExtensionAttribute9 -Replace @{Description="$newDesc - Ajouté aux exception Infomut"}
        Write-Success "User $($user.SamAccountName) sera affecté par le processus infomut"
    }
}

Function Set-ExpirationDates(){
    Set-InfomutExpirationDates
}

Function Set-InfomutExpirationDates(){
    Write-Prompt "Séléctionner l'export EZV" -noNewLine:$false
    $file = Get-FilePathFromDialog
    if($null -eq $file){
        Write-Err "Veuillez séléctionner un fichier" -noNewLine:$true
        return    
    }
    
    $users = Make-LeaversFromCSV $file
    Write-Host ""#On ajoute une ligne vide d'espacement entre les infos et les résultats
	foreach($user in $users){
		if($null -eq $user -or $null -eq $user.ADUser){
			continue
		}
		$isException = $user.ADUser.ExtensionAttribute9 -like $EXCEPTION_KEYWORD
		if($isException){
			Write-Info "User $($user.ADUser.SamAccountName) est dans la liste d'exceptions"
		}else{
			Process-InfomutUser $user
			Create-InfomutLogFile $user
		}
	}
	
    Write-Success "Traité $($users.Length) utilisateurs"
}

Function Set-SingleExpirationDate([string]$sam, [string]$date, [string]$ticketNumber){
    $username = Assert-StrArg $sam $_ "Nom d'utilisateur"
    $user = Get-ADUser $username -Properties ExtensionAttribute9, DistinguishedName
    if($user.ExtensionAttribute9 -eq $EXCEPTION_KEYWORD){
        Write-Error "L'utilisateur a été placé dans la liste d'exceptions pour les infomuts.`nMerci de le retirer via Set-Exception, puis de relancer cette fonction"
        return
    }
    $parsedDate = [DateTime]::MinValue
    if($date.Length -eq 0 -or ([DateTime]::TryParse($date, [ref]$parsedDate) -eq $false)){
        $parsedDate = Prompt-ForDate "Date d'expiration"
    }
    $ticketNumber = Assert-StrArg $ticketNumber '' "Numéro de ticket"
    $leaver = [Leaver]::new($user, $parsedDate, $ticketNumber)
    try{
        Process-InfomutUser $leaver
        Create-InfomutLogFile $leaver
    }catch{
        Write-Err "Erreur lors du traitement de $($user.SamAccountName)"
        Write-Err $_
    }
}

Function Process-ExpiringAccounts{ 
    $users = Get-ExpiredInfomutUsers $DAYS_BEFORE_NETTOYAGE
    $currDateString = (Get-Date -Format 'dd/MM/yy').ToString()
    $users | % { Set-ADUser $_.SamAccountName -Enabled:$false `
                    -Description "$((Get-Date -format 'dd/MM/yy').ToString()) Infomut J+$DAYS_BEFORE_NETTOYAGE, compte déplacé dans l'OU 'Nettoyage'"`
                    -Replace @{ExtensionAttribute12 = $currDateString}
                 Move-ADObject -Identity $_.DistinguishedName -TargetPath $NETTOYAGE_OU
                 Write-Success "Traité $($_.SamAccountName) avec succès"
                 "$($currDateString) J+$($DAYS_BEFORE_NETTOYAGE), déplacé dans l'OU $($NETTOYAGE_OU)" >> "C:\Infomut\$($_.SamAccountName).log"
    }
    $users | Export-Csv "C:\Infomut\Daily\$(Get-Date -Format 'yy/MM/dd')-Expiring.csv" -UseCulture -Encoding UTF8 -NoTypeInformation
}

Function Process-LeaverAccounts {
    $users = Get-LeaverAccounts
    $users | % { 
     Write-Prompt ("`n" + $_.SamAccountName) -noNewLine:$false
     $exportBasePath = "C:\Infomut\$($_.SamAccountName)"
     Export-UserGroups $_.SamAccountName "$exportBasePath.groups"
     Write-Success "Backup des groupes OK ($exportBasePath.groups)"
     Export-UserPersonalInfo $_.SamAccountName "$exportBasePath.info"
     Write-Success "Backup des infos personnelles OK ($exportBasePath.info)"

     Remove-AllGroups $_.SamAccountName
     Write-Success "User retiré de tout les groupes"

     Clear-UserPersonalData $_.SamAccountName
     Write-Success "Supprimé infos personnelles"

     Set-ADUser -Identity $_.SamAccountName -Description "$(Get-Date -format 'dd/MM/yyyy')"
     
     Move-ADObject -Identity $_ -TargetPath $LEAVERS_OU
     Write-Success "Déplacé dans l'OU Leaver"

     "$(Get-Date -Format 'dd/MM/yy') J+$($DAYS_BEFORE_LEAVERS), déplacé dans Leavers" >> "C:\Infomut\$($_.SamAccountName).log"
    }
    $users | Export-Csv "C:\Infomut\Daily\$(Get-Date -Format 'yy/MM/dd')-Leavers.csv" -UseCulture -Encoding UTF8 -NoTypeInformation
}

Function Export-UserGroups([string]$sam, [string]$path){
    $sam = Assert-StrArg $sam '' "Nom d'utilisateur"
    $path = Assert-StrArg $path '' "Emplacement de l'export"
    Get-ADUser -Identity $sam -Properties MemberOf | Select -ExpandProperty MemberOf > $path
}

Function Import-UserGroups([string]$sam, [string]$path, [bool]$silent = $false){
    $sam = Assert-StrArg $sam '' "Nom d'utilisateur"
    $path = Assert-StrArg $path '' "Emplacement de l'export"

    Get-Content $path | % { try {
      $group = $_
      Add-ADGroupMember -Identity $group -Members $sam ;
      if($silent -eq $false) {Write-Success "Ajouté $(Get-CNFromDistinguishedName $_)"} }
     catch{
       if($silent -eq $false) {Write-Err "Erreur lors de l'ajout du groupe $group";Write-Err $_} } }
}

Function Export-UserPersonalInfo([string]$sam, [string]$path){
    $sam = Assert-StrArg $sam '' "Nom d'utilisateur"
    $path = Assert-StrArg $path '' "Emplacement de l'export"
    Get-ADUser -Identity $sam -Properties EmployeeID, EmployeeType, Manager, ExtensionAttribute10, ExtensionAttribute14, TelephoneNumber, HomePhone, Mobile, IpPhone, OfficePhone |`
     Select EmployeeID, EmployeeType, Manager, ExtensionAttribute10, ExtensionAttribute14, TelephoneNumber, HomePhone, Mobile, IpPhone, OfficePhone |`
     Export-Csv -Path $path -Encoding UTF8 -NoTypeInformation
}

Function Import-UserPersonalInfo([string]$sam, [string]$path, [bool]$silent = $false){
    $sam = Assert-StrArg $sam '' "Nom d'utilisateur"
    $path = Assert-StrArg $path '' "Emplacement de l'export"

    $result = @{}
    $data = Import-Csv $path

    $data.PSObject.Properties | ? {$null -ne $_.Value -and $_.Value.Length -gt 0} | % { $result.Add($_.Name, $_.Value) }
    if($result.Count -eq 0){
        return
    }

    Set-ADUser $sam -Replace $result
    if($silent -eq $false){
        Write-Success "Importé les données suivante :"
        $result | FL
    }
    
}

Function Restore-LeaverOriginalOU([string]$sam, [bool]$silent = $false){
    $user = Get-ADUser $sam -Properties DistinguishedName, EmployeeType
    if($user.EmployeeType -eq 1){
        $destinationOU = $INTERNAL_OU
    }else{
        $destinationOU = $EXTERNAL_OU    
    }

    Move-ADObject -Identity $user.DistinguishedName -TargetPath $destinationOU
    if($silent -eq $false){
        Write-Success "Déplacé l'utilisateur dans l'OU $oldOU"
    }
}

Function Rollback-Infomut([string]$sam){
begin{
    $sam = Assert-StrArg $sam $_ "`nNom d'utilisateur"
    $user = Get-ADUserByUsername $sam
    $basePath = "C:\Infomut\$($user.SamAccountName)"
    $checklist = [Checklist]::new("Rollback infomutations", 
        ("Import des infos personnelles", "Import des groupes", "Réactivation du compte AD", "Déplacement dans l'OU"))
}
process{
    $checklist.Start()

    if([System.IO.File]::Exists("$basePath.info") -eq $false){
        $checklist.FatalError()
        Write-Err "Pas trouvé l'export des données personnelles ($basePath.info)" -noNewLine:$true
        return
    }
    if([System.IO.File]::Exists("$basePath.groups") -eq $false){
        $checklist.FatalError()
        Write-Err "Pas trouvé l'export des groupes ($basePath.groups)" -noNewLine:$true
        return
    }

    try{
        Import-UserPersonalInfo $user.SamAccountName -path "$basePath.info" -silent $true
        $checklist.SetStateAndGoToNext($true)
    }catch{
        $checklist.FatalError()
        Write-Err "Erreur lors de l'ajout des infos personnelles" -noNewLine:$true
        Write-Err $_
        return
    }

    try{
        Import-UserGroups $user.SamAccountName -path "$basePath.groups" -silent $true
        $checklist.SetStateAndGoToNext($true)
    }catch{
        $checklist.FatalError()
        Write-Err "Erreur lors de l'ajout des groupes" -noNewLine:$true
        Write-Err $_
        return
    }


    try{
        Set-ADUser -Identity $user.DistinguishedName -Enabled:$true -Clear ("ExtensionAttribute11", "ExtensionAttribute12")
        $checklist.SetStateAndGoToNext($true)
    }catch{
        $checklist.FatalError()
        Write-Err "Erreur lors de la réactivation du comtpe" -noNewLine:$true
        Write-Err $_
        return
    }

    try{
        Restore-LeaverOriginalOU $user.SamAccountName -silent:$true
        $checklist.SetState($true)
    }catch{
        $checklist.FatalError()
        Write-Err "Erreur lors du changement d'OU" -noNewLine:$true
        Write-Err $_
        return
    }
}
    
}

Function Get-InfomutExceptionList {
    return (Get-ADUser -Filter "ExtensionAttribute9 -eq 'NoLeavers'" -Properties AccountExpirationDate | Select Enabled, AccountExpirationDate,  Name, GivenName, Description)
}

Function Get-InfomutProgress([string]$username){
    $username = Assert-StrArg $user $_ "Nom d'utilisateur"
    $user = Get-ADUser $username -Properties ExtensionAttribute11, ExtensionAttribute12

    if($user.ExtensionAttribute11.Length -eq 0){
        Write-Prompt "L'utilisateur n'est pas encore dans le processus infomut"
    }elseif($user.ExtensionAttribute12.Length -eq 0){
        Write-Prompt "L'utilisateur est à la première étape, son compte n'a pas encore expiré, ou il n'a pas passé la période de grâce avant le déplacement dans l'OU Nettoyage"
    }elseif($user.DistinguishedName.Contains("OU=13-Nettoyage")){ #On regarde si l'utilisateur est dans l'OU "Nettoyage"
            Write-Prompt "L'utilisateur est à l'étape 'Nettoyage' (J+3), il possède toujours une licence et sa boîte mail sont toujours actif"
    }elseif($user.DistinguishedName.Contains("OU=Leavers")){
            Write-Prompt "L'utilisateur est à l'étape 'Leavers' (J+7), sa boîte mail a été désactivée, et sa licence a été retirée"
    }
}
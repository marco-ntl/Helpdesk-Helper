using module RTS-Components #Nécessaire pour l'objet Checklist
$LICENSE_GROUP = "RTS-G-L-M365_E3"
$RTS_OU = 'OU=RTS,OU=Units,DC=media,DC=int'
$RTS_USERS_OU = "OU=Users,$RTS_OU"

Function Add-UserToMailbox([string]$mailbox, $users, [boolean]$silent = $false) {
    begin {
        $mailbox = Assert-StrArg $mailbox $_ "Adresse mail ou alias de la boîte partagée"
        $users = Assert-StrArg $users $null "Username ou email de la personne"

        if($null -ne $users -and ($users.GetType() -eq [String] -and $users.Length -gt 0)){
            #La fonction traite des tableaux d'username, donc on créé un tableau qui ne contient que l'username
            $users = @(Get-ADUserByUsername $users)
        }elseif($users.GetType().BaseType -eq [System.Array]){
            $users = $users | % { Get-ADUserByUsername $_ }
        }else{
            $users = (Read-HostMultiline "Membres (1 par ligne):`n" | % { Get-ADUserByUsername $_ })
        }
        $items = $("Connexion Office365", "Ajout du droit 'Send As'", "Ajout du droit 'Full Access'")
        $checkList = [checklist]::new("Ajout des droits à une BAL partagée", $items)
        $checkList.SetSilent($silent)

    }
    Process {
        
        $checkList.Start() #"Connexion à O365"

        Import-O365Session -silent:$true
        $checkList.SetStateAndGoToNext($true) #Passe de "Connexion à O365" à "Send As"

        try {
            $users | % { Add-RecipientPermission -Identity $mailbox -Trustee $_.UserPrincipalName -AccessRights SendAs -Confirm:$false -WarningAction:SilentlyContinue -ErrorAction:Stop } | Out-Null
        }
        catch {
            $checklist.FatalError()
            Write-Err "Erreur lors de l'ajout des droits 'Send As', merci de vérifier les paramètres..."
            Write-Err $_
            return
        }
        $checklist.SetStateAndGoToNext($true) #Send As -> Full acces

        try {
            $users | % { Add-MailboxPermission -Identity $mailbox -User $_.UserPrincipalName -AccessRights FullAccess -InheritanceType All -Confirm:$false -WarningAction:SilentlyContinue -ErrorAction:Stop } | Out-Null
        }
        catch {
            $checklist.FatalError()
            Write-Err "Erreur lors de l'ajout des droits 'Full Access'..."
            Write-Err $_
            return
        }
        $checklist.SetState($true)

    }
    End {
        while ($silent -eq $false -and $val -ne $false) {
            #Si $session n'est pas nul, c'est que la fonction a été appelée par une autre fonction, et on ne devrait pas interagir avec l'utilisateur
            $val = (Ask-Redo -choices ('Boîte', 'Personne')).ToLower()
            if ($val -ne $false) {
                if ($val -eq 'a') {
                    Add-UserToMailbox -silent:$true
                }
                elseif ($val -eq 'b') {
                    Add-UserToMailbox -mailbox $mailbox -silent:$true
                }
                else {
                    Add-UserToMailbox -user $users -silent:$true
                }
            }
        }
    }

}

Function Enable-MailboxO365([string]$lastName, [string]$firstName, [bool]$notification = $false) {
    begin {
        $lastName = Remove-Accents (Assert-StrArg $lastName $_ "Nom de la personne")
        $firstName = Remove-Accents (Assert-StrArg $firstName '' "Prénom de la personne")

        $sam = Format-SAM  $lastName $firstName
        
        $Emailaddress = ($firstName + '.' + $lastName + '@rts.ch')
        $checklist = [Checklist]::new("Paramètrage Office 365", $(
                "Connexion à Office 365"
                "Activation de la Mailbox";
                "Paramètrage du calendrier";
                "Ajout de la licence Office 365";
            ))
    }
    process {
        $checklist.Start()
        Import-O365Session $true
        $checklist.SetStateAndGoToNext($true)
        try{
            Set-Mailbox -Identity $emailaddress -SingleItemRecoveryEnabled $true -AuditEnabled $true `
                -AuditDelegate "Update, Move, MoveToDeletedItems, SoftDelete, HardDelete, SendAs, SendOnBehalf, Create" -AuditOwner "Move, MoveToDeletedItems, SoftDelete, HardDelete"`
                -RetentionPolicy "SRG SSR" -ErrorAction Stop | Out-Null
        }catch{
            $checklist.FatalError()
            Write-Err "Erreur lors de l'attribution des paramètres de base, veuillez vérifier l'adresse mail..."
            Write-Err $_
            return
        }

        $checklist.SetStateAndGoToNext($true)
        try{
        #Paramètrage du calendrier
        Set-MailboxCalendarConfiguration -Identity $Emailaddress -WeekStartDay Monday -WorkDays Weekdays -WorkingHoursStartTime 08:00:00 -WorkingHoursEndTime 17:00:00 -WorkingHoursTimeZone "W. Europe Standard Time" -ShowWeekNumbers $True
        #Disable ReplyToAll in OWA
        Set-MailboxMessageConfiguration -Identity $Emailaddress -IsReplyAllTheDefaultResponse $False
        #Disable Clutter
        Set-Clutter -Identity $Emailaddress -Enable $false | Out-Null
        }catch{
            $checklist.FatalError()
            Write-Err "Erreur lors de l'attribution des paramètres du calendrier, fermeture..."
            Write-Err $_
            return
        }
        $checklist.SetStateAndGoToNext($true)

        #Ajout de la licence
        try{
            Add-ADGroupMember -Identity $LICENSE_GROUP -Members $sam -ErrorAction Stop | Out-Null
        }catch{
            $checklist.FatalError()
            Write-Err "Erreur lors de l'attribution de la licence, fermeture..."
            Write-Err $_
            return
        }
        $checklist.SetState($true)
        Write-Success "Boîte mail créée"
        if ($notification) {
            Show-Notification -ToastTitle "Boîte mail activée" -ToastText "La boîte mail de $lastName, $firstName a bien été activée"
        }
    }
}

Function New-RTSSharedMailbox([string]$name, [string]$description) {
    Begin {
        $name = Assert-StrArg $name $_ "Nom de la boite mail"
        if($name.Contains('@')){
            $name = $name.Substring(0, $name.IndexOf('@'))
            Write-Info "Le nom ne peut pas contenir le charactère '@', il sera tronqué"
            Write-Info "Nouveau nom : $name"
        }
        $formattedName = Format-SAM $name -isUser:$false
        $OU = "OU=10-Mailbox,OU=Users,OU=RTS,OU=Units,DC=media,DC=int"
        $description = Assert-StrArg $description ''  "Veuillez entrer une Description du style 'Boite mail pour l'émission X ou Y'"
        if(Get-ADUser -filter "SamAccountName -eq '$formattedName'"){
            do{
                Write-Err  "`nUn compte '$formattedName' existe déjà... Veuillez modifier votre saisie ou vérifier l'AD"
                $name = read-host (Write-Prompt "Username de la boite mail")
                $formattedName = Format-SAM $name -isUser:$false
            }While (Get-ADUser -filter "SamAccountName -eq '$formattedName'")
            $userPrinc = (Remove-Accents $name | Format-Mail) + '@rts.ch'
            if(Ask-Confirmation "Utiliser $userPrinc comme adresse mail ?" -eq $false){
                $userPrinc = ((Read-Host (Write-Prompt "Adresse mail")) | Remove-Accents | Format-Mail) + "@rts.ch"
            }
        }

        $DisplayName = $name + ' (RTS)'
        $userPrinc = (Remove-Accents $name | Format-Mail) + '@rts.ch'

        $checklist = [checklist]::new("Création de boîte aux lettres partagée", $(
                "Connexion au CAS",
                "Création de l'objet AD",
                "Réplication sur le CAS",
                "Activation de la boîte"
            ))
    }
    process {
        $checklist.Start() #Connexion CAS
        Get-CASSession -silent:$true
        $checklist.SetStateAndGoToNext($true) #Création objet ad
        #Write-Info "Création de l'objet AD Mailbox partagée en cours...`n"
        New-ADUser -Name $formattedName -UserPrincipalName $userPrinc -DisplayName $DisplayName -path $OU -Description $description -EmailAddress $userPrinc | Out-Null

        $checklist.SetStateAndGoToNext($true) #Réplication CAS
        #Write-Info "Mailbox créé, en attente de synchro..."
        Wait-ProgressBar 50 "Réplication sur le CAS"#Enlevé le sleep, à voir si ça pose problème
        $checklist.SetStateAndGoToNext($true) #Activation de la boîte
        #Write-Info "Synchro OK, activation de la boîte mail..."
        Send-CASRequest -ScriptBlock {param($name, $userPrinc) Enable-RemoteMailbox -Shared -Identity $name -RemoteRoutingAddress "$($name)@SRGSSR.mail.onmicrosoft.com" -PrimarySmtpAddress $userPrinc -alias $name } -argumentList $formattedName, $userPrinc | Out-Null
        #Get-Mailbox $name | Format-List PrimarySmtpAddress
        Send-CASRequest -ScriptBlock {param($name) Set-User -Identity $name -Company 'Radio Télévision Suisse' } -argumentList $formattedName | Out-Null
        $checklist.SetState($true)
        Write-Success "La boite aux lettres partagée '$name' vient d'être créée."
        Write-Info "Merci d'attendre $((Get-Date).AddHours(1).ToString('HH:mm')) que la migration se fasse automatiquement puis d'ajouter les membres en Full acc�s + Send As Permissions dans O365 exchange console"

        $mailData = (
            ("Nom de la boîte", $userPrinc),
            ("Boîte créée par", $env:USERNAME)
        )
        $mailHeader = "Boîte $userPrinc créée"
        Send-MailFromTemplate -subject $mailHeader -from 'Postmaster-RTS@rts.ch' -to 'Postmaster-RTS@rts.ch' -header $mailHeader -tableData $mailData
    }
}

Function New-RTSDistributionGroup([string]$name, [string]$manager, [Object[]]$members, $silent = $false) {
    Begin {
        $fullName = Assert-StrArg $name $_ "Nom de la liste"
        $sam = Format-SAM $fullName -isUser:$false
        $name = Format-Mail $fullName

        While ($null -ne $(Get-ADGroup -Filter "SamAccountName -eq '$sam' -or mail -eq '$name@rts.ch'")) {
            Write-Err "Un compte avec le SAM '$sam' ou l'adresse '$name@rts.ch' existe déjà"
            if(Ask-Confirmation "Modifier seulement le SAM ? ('Non' si l'email est déjà utilisé)"){
                Write-Prompt "SAM: "
                $sam = Format-SAM (Read-Host) -isUser:$false
            }else{
                Write-Prompt "Merci d'entrer un autre nom: "
                $fullName = Read-Host 
                $sam = Format-SAM $fullName -isUser:$false
                $name = Format-Mail $fullName

            }
        }
        
        $DispName = $fullName + " (RTS)"

        $manager = Assert-StrArg $manager '' "Nom du propriétaire de la liste (login ou adresse mail)" #@TODO Remplacer liste par multi-line read-host
        $manager = (Get-ADUserByUsername $manager).UserPrincipalName #On utilise get-aduserbyusername pour s'assurer que l'utilisateur existe
        if ($null -eq $members) {
            $members = Read-HostMultiline "Membres (1 par ligne):`n" | % { Get-ADUserByUsername $_ } | % { return $_.UserPrincipalName } 
            #While (Ask-Confirmation "`nAjouter un membre ?"){
            #    Write-Prompt "Veuillez entrer l'username de la personne: " -NoNewline:$true
            #    $members += (Get-ADUserByUsername (Read-Host)).UserPrincipalName #On utilise get-aduserbyusername pour s'assurer que l'username est valide
            #}
        }
        $OU = "OU=Distribution Groups,OU=Messaging,OU=Groups,OU=RTS,OU=Units,DC=media,DC=int"
        $checklist = [checklist]::new("Création d'une liste de distribution", $(
                "Connexion au CAS",
                "Création de la liste",
                "Ajout des membres"
            ))
        $checklist.SetSilent($silent)
    }
    Process {
        $checklist.Start() #Connexion au CAS
        Get-CASSession -silent:$true
        if ($null -eq (Send-CASRequest -ScriptBlock {param($sam) Get-DistributionGroup $sam -erroraction SilentlyContinue } -argumentList $sam)) {
            #TODO implémenter REDO
            $checklist.SetStateAndGoToNext($true)#Création de la liste
            
            #Write-Info "La liste de Distribution $nomDistribGroup n'existe pas, création en cours..."
            Send-CASRequest -ScriptBlock {param($name, $DispName, $OU, $sam) New-DistributionGroup -Name $name -DisplayName $DispName -OrganizationalUnit $OU -SamAccountName $sam -MemberJoinRestriction 'Closed' -MemberDepartRestriction 'Closed'  -Alias $name -Type Distribution -erroraction SilentlyContinue } -argumentList $name,$DispName,$OU, $sam | Out-Null
            Send-CASRequest -ScriptBlock {param($sam,$Manager) Set-DistributionGroup -Identity $sam -ManagedBy $Manager -RequireSenderAuthenticationEnabled $false -erroraction SilentlyContinue } -argumentList $sam,$Manager | Out-Null
            $checklist.SetStateAndGoToNext($true)
            #Write-Info "Ajout des membres..."
            $members | % { Send-CASRequest -ScriptBlock {param($sam,$_) Add-DistributionGroupMember -Identity $sam -Member $_ -errorAction SilentlyContinue } -argumentList $sam,$_ } | Out-Null
            $checklist.SetState($true)
            if ($silent -eq $false) {
                Write-Success "Les membres ont bien été ajoutés`nLa liste $fullName a bien été créée."
            }
        }
        Else {
            Write-Err "La liste de Distribution $fullName existe déjà"
        }
    }
}

Function Rename-DistributionGroup([string]$groupName, [string]$newName) {
    Begin {
        $groupName = Assert-StrArg $groupName $_ "Nom de la liste"
        $newName = Assert-StrArg $newName '' "Nouveau nom de la liste"
        $checklist = [Checklist]::new("Renommer une liste de distribution", $(
                "Connexion au CAS", 
                "Création de la nouvelle liste",
                "Insertion dans les groupes de l'ancienne liste",
                "Suppression de l'ancienne liste"
            ))
    }
    Process {
        $checklist.Start()
        Get-CASSession -silent:$true
        $checklist.SetStateAndGoToNext($true)

        try {
            $group = Send-CASRequest -ScriptBlock {param($groupName) Get-DistributionGroup $groupName } -argumentList $groupName
        }
        catch {
            Write-Err "Pas trouvé de liste $groupName, fermeture..."
            exit    
        }
        $manager = $group.ManagedBy[0].Substring($group.ManagedBy[0].lastIndexOf('/') + 1) #ManagedBy contient le Path vers l'objet AD, on veut juste l'username
        $members = Send-CASRequest -ScriptBlock {param($group) Get-DistributionGroupMember $group.Identity } -argumentList $group | Select Name | % { $_.Name }
        New-RTSDistributionGroup -name $newName -manager $manager -members $members -silent:$true
        $checklist.SetStateAndGoToNext($true)#Insertion dans les groupes de l'ancienne liste

        Get-ADPrincipalGroupMembership $groupName | % { Add-ADGroupMember -Identity $_ -members $newName }
        $checklist.SetStateAndGoToNext($true)#Write-Info "Suppression de l'ancienne liste..."

        Send-CASRequest -ScriptBlock {param($group) Remove-DistributionGroup -Identity $group.Identity -Confirm:$false } -argumentList $group | Out-Null
        $checklist.SetState($true)
    }
}

Function Remove-UserFromMailbox([string]$user, [string]$mailbox, [bool]$silent = $false) {
    Begin {
        $user = Assert-StrArg $user $_ "Username ou adresse mail de l'utilisateur"
       
        $checkList = [Checklist]::new("Retirer user BAL partagée", $(
                "Connexion à Office 365",
                "Suppression du droit 'Send As'"
                "Suppression du droit 'Full access'"
            ))
        $checkList.SetSilent($silent)
    }
    Process {
        $checklist.Start() #Connexion à office 365
        Import-O365Session -silent:$true -forceImport:$true #On force l'import, car Remove-MailboxPermission est dispo dans ExchangeOnlineManagement et le CAS, mais ne fonctionnera que sur ExchangeOnlineManagement
        $checklist.SetStateAndGoToNext($true) #Suppression "Send As"

        try {
            Remove-RecipientPermission -Identity $mailbox -Trustee $user -AccessRights SendAs -Confirm:$false -WarningAction:SilentlyContinue #Warning action nécessaire, sinon un warning s'affiche
        }
        catch {
            $checklist.FatalError()
            Write-Err "Erreur lors de la suppression des droits, veuillez vérifier l'adresse mail..." #@TODO better error handling
            Write-Err $_
            return
        }
        $checklist.SetStateAndGoToNext($true) #Suppression "Full access"
        try {
            Remove-MailboxPermission -Identity $mailbox -User $user -AccessRights FullAccess -BypassMasterAccountSid -Confirm:$false -WarningAction:SilentlyContinue #Warning action nécessaire, sinon un avertissement va parfois s'afficher, ce qui casse l'affichage
        }
        catch {
            $checklist.FatalError()
            Write-Err "Erreur lors de la suppression des droits, veuillez vérifier l'adresse mail..." #@TODO better error handling
            Write-Err $_
            return
        }
        $checklist.SetState($true)
    }
    End {
        if (($silent -eq $false) -and (Ask-Confirmation "Recommencer ?")) {
            Remove-UserFromMailbox
            return
        }
    }
}

Function Set-MailboxHidden($mailbox, $hidden = $null, $isMailbox = $null, [bool]$OnPremises = $false) {
    begin{
        $mailbox = Assert-StrArg $mailbox $_ "Addresse mail"
        if($null -eq $hidden){
            $hidden = Ask-Confirmation "Cacher du carnet d'adresse"
        }
        if($null -eq $isMailbox){
            $isMailbox = Ask-Confirmation "Est-ce une boîte mail ?"
        }
        $checklist = [checklist]::new("Modification de la visibilité d'une BAL/DL", @(
        "Paramètrage de la liste d'adresse globale"
        ))
    }
    process{
        if($isMailbox -and $OnPremises -eq $false){
            $checklist.Inject("Connexion O365", 0)
            $checklist.Start()
            Import-O365Session -silent:$true
        }else{
            $checklist.Inject("Connexion CAS", 0)
            $checklist.Start()
            Get-CASSession -silent:$true
        }
        if($isMailbox){
            if($OnPremises){
                $mailboxID = (Send-CASRequest -scriptBlock {param($mailbox) Get-RemoteMailbox $mailbox } -argumentList $mailbox).Identity #Il faut explicitement prendre Identity, car sinon c'est le displayName qui sera récupéré
            }else{
                $mailboxID = Get-Mailbox $mailbox | Out-Null
            }
        }else{
            $mailboxID = (Send-CASRequest -ScriptBlock {param($mailbox) Get-DistributionGroup $mailbox } -argumentList $mailbox).PrimarySmtpAddress
        }
        if($isMailbox -eq $true -and $OnPremises -eq $false){
            $checklist.FatalError()
            Write-Err "Mailbox $mailbox introuvable, tentative depuis le CAS..."
            return Set-MailboxHidden -mailbox $mailbox -hidden $hidden -isMailbox $isMailbox -OnPremises:$true
        }elseif($null -eq $mailboxID){
            Write-Err "Mailbox $mailbox introuvable..."
            return
        }
        $checklist.SetStateAndGoToNext($true)
        try{
        if($isMailbox){
            if($OnPremises){
                Send-CASRequest -scriptBlock {param($mailboxID, $hidden) Set-RemoteMailbox $mailboxID -HiddenFromAddressListsEnabled:$hidden -WarningAction:SilentlyContinue } -argumentList $mailboxID, $hidden | Out-Null
            }else{
                Set-Mailbox $mailbox -HiddenFromAddressListsEnabled:$hidden -WarningAction:SilentlyContinue | Out-Null
            }
        }else{
            Send-CASRequest -ScriptBlock {param($mailbox,$hidden) Set-DistributionGroup $mailbox -HiddenFromAddressListsEnabled:$hidden -WarningAction:SilentlyContinue } -argumentList $mailboxID, $hidden | Out-Null
        }
        $checklist.SetState($true)
        }catch{
            $checklist.FatalError()
            Write-Err "Erreur lors de la modification, fermeture..."
            return
        }
        Write-Info "Les changements ont bien été effectués, ils devraient être effectifs sous 24h"
        return
    }
}

Function Set-DistributionListManagers($distributionList, $managers, $remove = $null){
process {
        do{
            $dlName = Assert-StrArg $distributionList $_ 'Nom de la liste'
            $distributionList = Get-AdGroup -Filter "Name -eq '$dlName'" -Properties ManagedBy, MsExchCoManagedByLink
        }while($null -eq $distributionList)
        
        $mainManager = $distributionList.ManagedBy 
        $currManagers = [System.Collections.ArrayList]@(($distributionList.msExchCoManagedByLink + $distributionList.ManagedBy) | ? { $null -ne $_ }) #ArrayList est nécessaire, car les arrays sont de taille fixe. Le tableau contient aussi un 'null', que l'on doit retirer

        if($currManagers.Count -eq 0){
            Write-Info "`nLa liste n'a pas de managers`n"
        }else{
            ([SimpleTable]::new("Managers $($distributionList.Name)", ($currManagers | % { Get-CNFromDistinguishedName $_ }))).Show()
        }

        if($null -eq $remove){
            $remove = (Request-MultipleChoices "Gestion des propriétaires d'une DL" ("1 Ajouter", "2 Supprimer") -eq 0)
        }
        
        if($remove){
            if($distributionList.ManagedBy.Length -eq 0 -and $distributionList.msExchCoManagedByLink.Count -eq 0){
                Write-Err "La liste n'a pas de manager"
                return
            }
            if($managers -eq $null){
                $managers = [System.Collections.ArrayList]::new()
        
                do{
                    $select = [SimpleSelect]::new("Managers à retirer", ($currManagers | % { Get-CNFromDistinguishedName $_ }))
        
                    $managerIndex = $select.AskUser()
                    $managers.Add($currManagers[$managerIndex]) | Out-Null #$arrayList.Add retourne '0', qui est écrit dans la console si on retire le out-null
                    $currManagers.RemoveAt($managerIndex) | Out-null #Idem
                    if($currManagers.Count -le 0){
                        Write-Info "Plus de managers`n"
                        break 
                    }
                }while(Ask-Redo -prompt "Retirer quelqu'un d'autre ?" -eq $true)
            }
            if($managers.Contains($mainManager)){
                Set-ADGroup $distributionList -Clear ManagedBy
                $managers.Remove($mainManager)
            }
            $managers | % { Set-ADGroup $distributionList -Remove @{msExchCoManagedByLink = $_} }
        }else{
            if($null -eq $managers){
                Write-Prompt "Propriétaires (1 par ligne):`n"
                $managers = Read-HostMultiline | % { Get-ADUserByUsername $_ }
                }
                $managers | % { Set-ADGroup $distributionList -Add @{msExchCoManagedByLink = $_.DistinguishedName} }
            }
    
    }
    end{
        if(Ask-Redo){
            Write-Host "`n" -NoNewline:$true
            Set-DistributionListManagers
        }
    }

}
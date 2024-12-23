using module RTS-Components #Nécessaire pour l'objet Checklist

$DEFAULT_PWD = 'DEFAULTPWD'

$RTS_OU = 'OU=RTS,OU=Units,DC=media,DC=int'
$RTS_USERS_OU = "OU=Users,$RTS_OU"
$INTERNAL_OU = "OU=2_Current, OU=Internal,$RTS_USERS_OU"
$EXTERNAL_OU = "OU=2_Current, OU=External,$RTS_USERS_OU"


$GENERAL_GROUPS = $("RTS-G-CRA-CONNECT-USERS", "RTS-G-DomainUsers", "MEDIA-G-DomainUsers", "RTS-G-CLOUDPRX", "RTS-G-SMS-Auth", "RTS-G-O365MFAenable")
$INTERNAL_GROUPS = $($GENERAL_GROUPS) + $("RTS-G-TS-Publ_Internal_Users-RTS", "RTS-Mail-Users-Int")
$EXTERNAL_GROUPS = $($GENERAL_GROUPS) + $("RTS-G-TS-Publ_External_Users-RTS", "RTS-Mail-Users-Ext")

$POST_MASTER_DISTRIB = "RTS-OP-UT-MDD-PostMaster@rts.ch"

Function Add-DefaultGroups([string]$userName, [bool]$isInternal) {
    process {
        $user = Assert-StrArg $userName $_ 'Username'
        if ($null -eq $isInternal) {
            $isInternal = Ask-Confirmation "Utilisateur interne ?"
        }
        if ($isInternal) {
            $groups = $INTERNAL_GROUPS
        }
        else {
            $groups = $EXTERNAL_GROUPS
        }
        $groups | ForEach-Object { Add-ADGroupMember -Identity $_ -Members $user | Out-Null }
    }
}

Function Enable-Skype4B {
    param(
        [Parameter(Mandatory = $false)][string]$lastName,
        [Parameter(Mandatory = $false)][string]$firstName,
        [Parameter(Mandatory = $false)][string]$sam,
        [Parameter(Mandatory = $false)]$enableMailbox = $null
    )
    Begin {
        $lastName = Assert-StrArg $lastName $_ 'Nom'
        $firstName = Assert-StrArg $firstName '' "Prénom"
        $sam = Assert-StrArg $sam '' "Username"
        if ($enableMailbox -eq $null) {
            $enableMailbox = Ask-Confirmation "Activer la mailbox ?"
        }
    }
    Process {
        $checklist = [Checklist]::new("Activation de Skype4Business", $(
                "Ouverture de la session à distance Lync";
                "Activation du compte";
                "Synchronisation";
                "Attribution des droits";
            ))
        $checklist.Start()
        try {
            $lyncOptions = New-PSSessionOption -SkipRevocationCheck -SkipCACheck -SkipCNCheck
            $LyncSession = New-PSSession -ConnectionUri https://ucwebsvc01.media.int/ocspowershell -SessionOption $lyncOptions -Authentication NegotiateWithImplicitCredential -erroraction Stop
            Import-PSSession $LyncSession -AllowClobber | Out-Null 
        }
        catch {
            $lyncOptions = New-PSSessionOption -SkipRevocationCheck -SkipCACheck -SkipCNCheck
            $LyncSession = New-PSSession -ConnectionUri https://ucwebsvc01.media.int/ocspowershell -SessionOption $lyncOptions -Authentication NegotiateWithImplicitCredential
            Import-PSSession $LyncSession -AllowClobber | Out-Null 
        }
        #Write-Info "`nSession à distance ouverte, en attente de synchro..."
        Wait-ProgressBar 10 "En attente de synchro"
        $checklist.SetStateAndGoToNext($true)
        #Write-Info "Synchro OK, activation du compte..."
        try {
            Enable-CsUser -Identity $sam -RegistrarPool "fepool01.media.int" -SipAddress "sip:$($(Remove-Accents $firstName)+'.'+$(Remove-Accents $lastName) +'@rts.ch')" -WarningAction:"SilentlyContinue" | Out-Null
        }
        catch {
            $checklist.FatalError()
            Write-Err "Erreur lors de l'activation du compte, fermeture..."
            Write-Err $_
            return
        }
        $checklist.SetStateAndGoToNext($true)
        
        $i = 0
        $txt = "En attente de réplication du compte..." 
        Write-Progress -Activity $txt -Status "20s maximum" -PercentComplete 0 #On initialise la barre de chargement
        while ($null -eq $(Get-CsUser -Filter "SamAccountName -eq '$sam'")) {
            #Tant que l'utilisateur n'est pas trouvé c'est que la synchro n'est pas faite
            Start-Sleep 1
            $i += 1
            $percentage = $i / 20 * 100
            if ($percentage -gt 100) {
                $percentage = 100
            }
            Write-Progress -Activity $txt -Status "Encore $(20 - $i)s maximum" -PercentComplete $percentage
        }
        Write-Progress -Activity $txt -Status "OK" -PercentComplete 100
        $checklist.SetStateAndGoToNext($true)
        #While ($null -eq $(Get-CsUser -Filter "SamAccountName -eq '$sam'")) {
        #@TODO FIX
        #    Start-Sleep 1 
        #}
        try {
            Get-Csuser $sam | Grant-CsConferencingPolicy -PolicyName "Enterprise CAL (Allow Conferencing)" -WarningAction:"SilentlyContinue"
        }
        catch {
            $checklist.FatalError()
            Write-Err "Erreur lors de l'ajout des droits de conférence, fermeture..."
            Write-Err $_
            return        
        }
        $checklist.SetState($true)
        Write-Success "Skype for Business est bien activé`n"
        #if(Ask-Confirmation "Activer la mailbox ?"){
        if ($enableMailbox) {
            $mail = (Get-ADUser -Identity $sam -Properties UserPrincipalName).UserPrincipalName
            Wait-UntilMailboxSynced $mail
            Enable-MailboxO365 $lastName $firstName -notification:$true
        }
        else {
            Write-Info "Merci d'attendre $((Get-Date).AddHours(1).ToString('HH:mm')) puis de lancer Enable-MailboxO365"
        }
        #Sympa, mais le script demande toujours les credentials à l'utilisateur donc useless pour l'instant
        #$taskAction = New-ScheduledTaskAction -Execute "Powershell.exe" -Argument "Enable-Mailbox-O365 -lastName $lastName -firstName $firstName -notification:$true"
        #$taskTrigger = New-ScheduledTaskTrigger -Once -At ([datetime]::Now).AddHours(1)
        #Register-ScheduledTask -TaskName "Enable-Mailbox_$lastName-$firstName" -User $(whoami) -Action $taskAction -Trigger $taskTrigger -TaskPath "Mediadesk" | Out-Null
        #Write-Info "Tâche planifiée créée, la boîte mail sera activée à $((Get-Date).AddHours(1).ToString('HH:mm'))"
    }
    End {
        Write-Info "Fermeture de la session à distance..."
        Remove-PSSession $LyncSession
    }
}

Function Enable-UserNoARS([string]$username, [string]$manager, [string]$employeeID, $expDate, $isInternal = $null) { 
    Begin {
        $username = Assert-StrArg $username $_ 'Username'
        $identity = Get-ADUserByUsername $username

        $manager = Assert-StrArg $manager '' "Username du manager"
        $managerIdentity = Get-ADUserByUsername $manager

        if ($null -eq $isInternal) {
            $isInternal = Ask-Confirmation 'Est-ce que le compte est interne ?'
        }

        if ($isInternal) {
            $employeeID = Assert-StrArg $employeeID '' "Matricule"
            $ou = $INTERNAL_OU
            $ouName = 'interne'
        }
        else {
            $employeeID = '9x9' + $username
            $ou = $EXTERNAL_OU
            $ouName = 'externe'

            $defaultExpDate = (Get-Date).AddYears(1)
            $expDate = Assert-StrArg $expDate '' "Merci d'entrer la date d'expiration (JJ/MM/AAAA) ($($defaultExpDate.ToString('dd/MM/yyyy')) par défaut)"
            if ($expDate.Length -le 0) {
                $expDate = $defaultExpDate
            }
            elseif ((Test-DateFormat -date $date) -eq $false) {
                $expDate = Prompt-ForDate -prompt "Format incorrect (JJ/MM/AAAA), réessayez"
            }
        }

        $checklist = [checklist]::new("Réactivation d'un compte en Leavers", @(
                "Réactivation du compte AD",
                "Insertion dans les groupes par défaut",
                "Reset du mot de passe ($DEFAULT_PWD)",
                "Définition des différents attributs",
                "Déplacement dans l'OU $ouName"
            ))

    }
    Process {
        $checklist.Start()#"Réactivation du compte AD",
        try {
            Enable-ADAccount -Identity $identity
        }
        catch {
            $checklist.FatalError()
            Write-Err "Erreur lors de l'activation du compte AD, fermeture..."
            Write-Err $_
            return
        }
        $checklist.SetState($true)
        
        try {
            if ($isInternal) {
                $checklist.Inject("Suppression de la date d'expiration", 0)
                $checklist.Next()
                Clear-ADAccountExpiration -Identity $identity
                $checklist.SetState($true)
            }
            else {
                $checklist.Inject("Assignation de la date d'expiration", 0)
                $checklist.Next()
                Set-ADAccountExpiration -Identity $identity -DateTime $expDate
                $checklist.SetState($true)
            }
        }
        catch {
            $checklist.FatalError()
            Write-Err "Erreur lors du paramètrage de la date d'expiration..."
            Write-Err $_
            return
        }

        $checklist.Next()#"Insertion dans les groupes par défaut",
        try {
            Add-DefaultGroups -username $username -isInternal $isInternal
        }
        catch {
            $checklist.FatalError()
            Write-Err "Erreur lors de l'ajout des groupes par défaut, fermeture..."
            Write-Err $_
            return
        }

        $checklist.SetStateAndGoToNext($true)#"Reset du mot de passe ($DEFAULT_PWD)",
        try {
            Set-ADAccountPassword -Identity $identity -Reset -NewPassword (ConvertTo-SecureString -String $DEFAULT_PWD -AsPlainText -Force)
        }
        catch {
            $checklist.FatalError()
            Write-Err "Erreur lors du paramètrage du mot de passe, fermeture..."
            Write-Err $_
            return
        }

        $checklist.SetStateAndGoToNext($true)#"Définition des différents attributs",
        try {
            Set-ADUser -Identity $identity -ChangePasswordAtLogon $true  -Replace @{extensionAttribute10 = $manager } -Manager $managerIdentity -EmployeeID $employeeID
        }
        catch {
            $checklist.FatalError()
            Write-Err "Erreur lors du paramètrage du compte, fermeture..."
            Write-Err $_
            return
        }

        $checklist.SetStateAndGoToNext($true)#"Déplacement dans l'OU $ouName"
        try {
            Move-ADObject -Identity $identity -TargetPath $ou #On move en dernier, car move un object change son identité (cn)
        }
        catch {
            $checklist.FatalError()
            Write-Err "Erreur lors du déplacement dans l'OU, fermeture..."
            Write-Err $_
            return
        }

        $checklist.SetState($true)
        Write-Info "Le compte a été réactivé, il faut cependant faire un ticket au national pour la réactivation de la boîte mail, si celle-ci a été désactivée"
    }
    end {
        if (Ask-Confirmation "Recommencer ?") {
            Enable-UserNoARS
        }
    }
}

class Template {
    static $SAM_TOKEN = "[SAM]"
    static $LASTNAME_TOKEN = "[LASTNAME]"
    static $FIRSTNAME_TOKEN = "[FIRSTNAME]"
    static $PWD_TOKEN = "[PWD]"
    static $WHOAMI_TOKEN = "[WHOAMI]"
}

Function New-User($isInternal = $null, [string]$lastName = $null, [string]$firstName = $null, [string]$title = $null, [string]$department = $null, $manager = $null, [string]$employeeID = $null, $expirationDate = $null, [string] $officePhone = $null, [string]$phone = $null, $enableSkypeAndMailbox = $null) {
    begin {
        if ($null -eq $isInternal) {
            $isInternal = Ask-Confirmation "Est-ce que l'utilisateur est interne ?"
            #do {
            #    #$type = (Read-Host "Type de compte ($(Underline('I'))nterne ou $(Underline('E'))xterne)").ToLower() Underline utilise les séquences d'échappement ANSI, qui ne sont pas compatibles avec toutes les version de PS
            #    $type = (Read-Host "Type de compte ([I]nterne ou [E]xterne)").ToLower() 
            #} While ($type[0] -ne 'i' -and $type[0] -ne 'e')
            #$isInternal = ($type[0] -eq 'i')
        }

        $lastName = Assert-StrArg $lastName '' "Nom"
        $firstName = Assert-StrArg $firstName '' "Prénom"
        $sam = Format-SAM -nom $lastName -prenom $firstName

        $title = Assert-StrArg $title '' "Fonction"
        $department = Assert-StrArg $department '' "Département" #@IDEA Liste de départements ?
        #$manager = Get-ADUserByUsername (Assert-StrArg $manager '' "Responsable (username)") -promptUserIfNotFound:$true
        $manager = Assert-StrArg $manager '' "Responsable (username)" #Certains comptes (eg. prestataire) n'ont pas de manager
        if($manager.Length -gt 0){
            $manager = Get-ADUserByUsername $manager
        }
        if ($isInternal -eq $true) {
            $employeeID = Assert-StrArg $employeeID '' "EmployeeID de la personne"
            $OU = $INTERNAL_OU
        }
        else {
            $employeeID = "9x9$($sam)"
            $OU = $EXTERNAL_OU

            $defaultExpDate = (Get-Date).AddYears(1)

            if ($null -eq $expirationDate -or $expirationDate.Length -le 0) {
                $expirationDate = Prompt-ForDate "Date de fin de contrat (JJ.MM.AAAA)($($defaultExpDate.ToString('dd.MM.yyyy')) par défaut)" -canBeNull:$true
            }
            else {
                $expirationDate = [datetime]::ParseExact($expirationDate, 'dd.MM.yy', [CultureInfo]::CurrentCulture).AddDays(1) #Si l'on donne une date sans spécifier d'heure, windows va set la date d'expiration au jour d'avant à 23h59
            }

            if ($expirationDate.Length -le 0) {
                $expirationDate = $defaultExpDate
            }
        }

        $OfficePhone = Assert-StrArg $officePhone '' "Numéro de téléphone interne (+41 58 236 xx xx)"
        $phone = Assert-StrArg $phone '' "Numéro de téléphone portable (+41 79 xxx xx xx)"
        if ($null -eq $enableSkypeAndMailbox) {
            $enableSkypeAndMailbox = Ask-Confirmation "`nActiver Skype4B et mailbox ?"
        }
        $Password = ConvertTo-SecureString -String $DEFAULT_PWD -AsPlainText -Force
        $DisplayName = $lastName + ', ' + $firstName + ' (RTS)'
    
        $userPrinc = $(Remove-Accents $firstName) + '.' + $(Remove-Accents $lastName) + '@rts.ch'

        While ($null -ne $(Get-ADUser -filter "SamAccountName -eq '$sam'")) {
            Write-Err  "`n$sam existe déjà... Veuillez saisir un autre login name"
            Write-Prompt "Login name: "
            $sam = Read-Host
        }
    }
    process {
        $checklist = [Checklist]::new("Création du compte", $(
                "Création du compte";
                "Synchro AD";
                "Ajout des groupes";
                "Ajout attributs généraux";
                "Connexion au CAS";
                "Création de la mailbox"
            ))
        $checklist.Start() #Création du compte
        try {
            new-ADUser -SamAccountName $sam -name $sam -surname $lastName -UserPrincipalName $userPrinc -DisplayName $DisplayName `
                -GivenName $firstName -Title $title -Department $department -Manager $manager.samAccountName -AccountPassword $Password `
                -Path $OU -Company 'Radio Télévision Suisse' -EmployeeID $employeeID -MobilePhone $phone `
                -OfficePhone $OfficePhone -ChangePasswordAtLogon:$true -Enabled:$True
            $checklist.SetStateAndGoToNext($true) #Synchro AD
        }
        catch {
            $checklist.FatalError()
            Write-Err "Erreur lors de la création du compte, fermeture..."
            Write-Err $_
            return
        }

        While ($null -eq $(Get-ADUser -Filter "SamAccountName -eq '$sam'" -SearchBase $RTS_USERS_OU)) {
            #Tant que l'utilisateur n'est pas trouvé c'est que la synchro n'est pas faite
            Start-Sleep 1 
        }
        $checkList.SetStateAndGoToNext($true) #Ajout des groupes

        try {
            $GENERAL_GROUPS | ForEach-Object { Add-ADGroupMember -Identity $_ -Members $sam | Out-Null }
        }
        catch {
            $checklist.FatalError()
            Write-Err "Erreur lors de l'ajout des groupes, fermeture..."
            Write-Err $_
            return
        }

        Wait-ProgressBar 5 "En attente de réplication" #Sinon l'user n'est parfois pas trouvé
        $checklist.SetStateAndGoToNext($true) #Ajout attributs généraux

        try {
            $attr = @{preferredLanguage = "fr-CH" }
            if ($null -ne $manager -and $manager.GetType().Name -ne "String") {
                $attr.Add('extensionAttribute10', $manager.samAccountName)
            }
            if ($null -ne $phone -and $phone.Length -gt 0) {
                $attr.Add('extensionAttribute14', $phone)
            }

            Set-ADUser -Identity $sam -Add $attr
        }
        catch {
            $checklist.FatalError()
            Write-Err "Erreur lors de l'ajout des groupes, fermeture..."
            Write-Err $_
            return
        }

        $checklist.SetState($true) #Pas de GoToNext car le prochain item doit être injecté selon si l'utilisateur est interne
                                   #Paramètres internes/eternes
        if ($isInternal -eq $false) {
            $checklist.Inject("Paramètres externes", 0)
            $checklist.Next()
            try {
                Set-ADAccountExpiration -Identity $sam -DateTime $expirationDate
                $EXTERNAL_GROUPS | ForEach-Object { Add-ADGroupMember -Identity $_ -Members $sam | Out-Null }
            }
            catch {
                $checklist.FatalError()
                Write-Err "Erreur lors de l'ajout des groupes, fermeture..."
                Write-Err $_
                return
            }

            $checklist.SetStateAndGoToNext($true)
        }
        else {
            $checklist.Inject("Paramètres internes", 0)
            $checklist.Next()
            try {
                $INTERNAL_GROUPS | ForEach-Object { Add-ADGroupMember -Identity $_ -Members $sam | Out-Null }
                Set-ADUSer -Identity $sam -Replace @{accountExpires = 0} #On set explicitement la date d'expiration à 0, sinon powershell lui donnera la valeur 9223372036854775807, qui pose un soucis avec Kaba
            }
            catch {
                $checklist.FatalError()
                Write-Err "Erreur lors de l'ajout des groupes, fermeture..."
                Write-Err $_
                return
            }
            $checklist.SetStateAndGoToNext($true)
        }
        
        try {
            Get-CASSession -silent:$true
            #Write-Info "En attente de synchro pour activation du lien exchange Onmicrosoft.com..."
            Wait-ProgressBar -time 20 -text "En attente de synchro pour activation du lien exchange Onmicrosoft.com..."
        }
        catch {
            $checklist.FatalError()
            Write-Err "Erreur lors de la connexion au CAS, fermeture..."
            Write-Err $_
            return
        }
        $checklist.SetStateAndGoToNext($true)
        #Start-Sleep 20 #Réduit le temps d'attente pour la synchro, à voir si cela pose problème
        #@TODO while try{get mailbox -errorAction Stop} pour éviter de hardcoder la sleep ?
        #While ($null -eq $(Get-RemoteMailbox -Filter { SamAccountName -eq $sam })) { #Impossible de faire Get-RemoteMailbox tant que celle-ci n'a pas été "enable"
        #    Start-Sleep 1 
        #}
        #Création dans cas.media.int du compte O365
        try {
            Send-CASRequest -ScriptBlock { param($sam) Enable-RemoteMailbox -Identity $sam -RemoteRoutingAddress "$sam@srgssr.mail.onmicrosoft.com" -Alias $sam } -argumentList $sam | Out-Null
        }
        catch {
            $checklist.FatalError()
            Write-Err "Erreur lors de la création de la boîte mail, fermeture..."
            Write-Err $_
            return
        }

        $checklist.SetState($true)
        Write-Success "`nLe compte AD a bien été créé ainsi que le lien @srgssr.mail.onmicrosoft.com"
        #Write-Success "Vous pouvez dès à présent activer Skype4B.`nIl faudra ensuite attendre 60 minutes et créer la Mailbox sur Exchange 365`n"

        #if (Ask-Confirmation "Activer Skype4B ?") {
        if ($enableSkypeAndMailbox) {
            Import-O365Session
            Wait-ProgressBar -time 10 -text "En attente de synchro avec Skype4B"
            Enable-Skype4B -lastName $lastName -firstName $firstName -sam $sam -enableMailbox:$true
        }

        $NEW_ACCOUNT_TEMPLATE = @(
            ("Compte créé", "$([Template]::LASTNAME_TOKEN), $([Template]::FIRSTNAME_TOKEN)"),
            ("Login", [Template]::SAM_TOKEN),
            ("Password", [Template]::PWD_TOKEN),
            ("Compte créé par", [Template]::WHOAMI_TOKEN)
        )
        $tokenValues = @(
            ([Template]::SAM_TOKEN, $sam),
            ([Template]::FIRSTNAME_TOKEN, $firstName),
            ([Template]::LASTNAME_TOKEN, $lastName),
            ([Template]::PWD_TOKEN, $DEFAULT_PWD),
            ([Template]::WHOAMI_TOKEN, $env:USERNAME)
        )

        $rawText = Format-StringFromTemplate $NEW_ACCOUNT_TEMPLATE $tokenValues
        $valuesDict = Make-DictionaryFromArray $rawText
        Write-Host #On écrit une ligne vide avant
        #$valuesDict | % { Write-Prompt ($_[0] + ': ');Write-Info $_[1]}
        foreach ($line in $valuesDict) {
            Write-Prompt ($line[0] + ': ')
            Write-Info $line[1]
        }
        Write-Host #On écrit une ligne vide après
        $subject = "Création du compte collaborateur $sam"
        $technicianMail = (Get-ADUser ($env:username.Substring($env:username.Indexof('-') + 1)) -Properties Mail).Mail
        Send-MailFromTemplate -subject $subject -from 'Postmaster-RTS@rts.ch' -to ($technicianMail, $POST_MASTER_DISTRIB) -header $subject -tableData $valuesDict
    }
    end {
        if (Ask-Confirmation "Recommencer ?") {
            New-User
        }
    }
 
}

Function Convert-ToGUID($uuid) {
    $uuid = Assert-StrArg $uuid $_ "Merci d'entrer l'UUID"
    return $uuid[6..7] + $uuid[4..5] + $uuid[2..3] + $uuid[0..1] + '-' + $uuid[10..11] + $uuid[8..9] + '-' + $uuid[14..15] + $uuid[12..13] + '-' + $uuid[16..19] + '-' + $uuid[20..31] -join ''
}




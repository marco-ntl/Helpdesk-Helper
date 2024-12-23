# Module RTS
#Création, désactivation et suppression des comptes utilisateurs RTS
#Création des BAL partagées
#
#Module à copier dans C:\Windows\System32\WindowsPowerShell\v1.0\Modules\RTS
#Faire Import-Module rts
#Modifié par Denis Ludovic le 06 mars 2018
#Mofdifié par Anwar Wasim le 09/11/2018 Ajout Manager dans extention10 pour service portal pour les externals users
#Modifié par Anwar Wasim le 18/03/2019 > Correction Bug ACL HomeDir
#Modifier par Anwar Wasim le 20/03/2019 #Masquer la BAL dans la GAL dans la fonction Desactivation_compte
#Mofdifié par Anwar Wasim le 27/06/2019 - ajout fr-CH dans l'attribut preferredLanguag=fr-CH
#Mofdifié par Anwar Wasim le 15/07/2019 - Desactivation_compte : Arrêt Suppression du Manager (nouvel processus pour Leavers)
#Mofdifié par Anwar Wasim le 27/07/2019 - Ajout 2 fonctions pour créations comptes Windows10 - interne et Externe
#Natalema - Ajouté déconnexion à la session Lync dans Enable-Skype4B pour ne plus avoir d'erreurs "too many concurrent shells"
#Natalema - Ajouté fonctions Format-name, Format-SAM et underline, retiré fonctions inutiles
#Natalema - Refactorisé les fonctions restantes pour supprimer le code en double
#Natalema - Réglé bug avec SIP lorsque le nom ou prénom contenait des accents
#Natalema - Transformé fonctions "w10-new-rts-interne" et "w10-new-rts-externe" en "New-User"
#Natalema - Ajouté des arguments à Enable-Skype4B. L'application demande désormais à l'utilisateur s'il souhaite activer Skype4B
#lorsqu'il créé le compte
#Natalema - Passé la fermeture des sessions à distance dans des blocks "End", pour éviter que la session en reste ouvert
#en cas de problème lors de l'exécution
#Natalema - Refactorisé New-DistributionGroup
#Natalema - Créée fonctions pour récupérer/importer les sessions à distance
#Natalema - Créé objet "Checklist", permettant de grandement faciliter et standardiser l'affichage des fonctions
#Natalema - Refactorisé "New-User" et "Enable-Mailbox-O365" pour utiliser les checklists
#Natalema - Ajouté système de cache pour les sessions. Désormais, le programme peut détecter si la session est déjà importée,
#et ne demandera les credentials plus qu'au premier lancement
#Natalema - Refactorisé [Checklist], certains calculs de taille étaient incorrects, causant des problèmes dans certains cas
#Natalema- Ajouté fonction New-RTS-ConferenceRoom
#Natalema - Ajouté "Wait-Until-Mailbox-Synced", qui permets de pauser l'exécution du script tant que la boîte mail spécifiée
#n'est pas visible dans O365. Cela permet d'activer une mailbox automatiquement, plutôt que d'attendre 1h puis de lancer "Enable-Mailbox"
#Natalema - Ajouté fonction "Enable-User-No-ARS" qui permet de réactiver un utilisateur en Leavers
#(la boîte mail doit quand même être réactivé par le national)
#Natalema - Ajouté fonction "Set-MailboxHidden" qui permet d'afficher/cacher une boîte aux lettres/liste de distribution du carnet d'adresse
#Natalema - Ajouté Error handling dans Set-MailboxHidden et Add-User-To-Mailbox
#Natalema - Refactorisé "Checklist" pour faire une base de framework d'interface
#Créé objet "SimpleTable" à partir de Checklist, l'idée étant d'avoir un composant simple sur lequel se baser pour les prochains
#La checklist hérite désormais de "SimpleTable", et redéfinit seulement les fonctions nécessaires (computelignlength, setstate, etc...)
#Créé objet "Simple Select", qui permet de demander un choix parmi une liste

# ╔═════════════════════════════════ TODO ══════════════════════════════════╗
# ║ Remplacer "Ask-Confirmation" par "Ask-Redo" quand possible              ║
# ║ Ajouter redo dans New-RTS-DistributionGroup si nom existe déjà          ║
# ║ Ajouter redo dans Rename-Distribution-Group si nom existe déjà          ║
# ╚═════════════════════════════════════════════════════════════════════════╝
using module RTS-Components #Nécessaire pour l'objet Checklist
Import-Module RTS-Infomut -WarningAction SilentlyContinue

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

$PROMPT_COLOR = [System.ConsoleColor]::Yellow
$script:CASSession = $null
$script:isO365Imported = $false
$REQUEST_ALL = "Tous"

$ALL_MODULES = (
                    ("AD", "E:\SCRIPTS\RTS-AD\RTS-AD", "C:\Windows\System32\WindowsPowerShell\v1.0\Modules\RTS-AD\RTS-AD"),
                    ("Components", "E:\SCRIPTS\RTS-Components\RTS-Components", "C:\Windows\System32\WindowsPowerShell\v1.0\Modules\RTS-Components\RTS-Components"),
                    ("Conference Room", "E:\SCRIPTS\RTS-ConferenceRooms\RTS-ConferenceRooms", "C:\Windows\System32\WindowsPowerShell\v1.0\Modules\RTS-ConferenceRooms\RTS-ConferenceRooms"),
                    ("Helpers", "E:\SCRIPTS\RTS-Helpers\RTS-Helpers", "C:\Windows\System32\WindowsPowerShell\v1.0\Modules\RTS-Helpers\RTS-Helpers"),
                    ("Exchange", "E:\SCRIPTS\RTS-Outlook\RTS-Outlook", "C:\Windows\System32\WindowsPowerShell\v1.0\Modules\RTS-Outlook\RTS-Outlook"),
                    ("Menu", "E:\SCRIPTS\RTS-Menu\RTS-Menu", "C:\Windows\System32\WindowsPowerShell\v1.0\Modules\RTS-Menu\RTS-Menu"),
                    ("Infomutation", "E:\SCRIPTS\RTS-Infomut\RTS-Infomut", "C:\Windows\System32\WindowsPowerShell\v1.0\Modules\RTS-Infomut\RTS-Infomut"),
                    ("Custom Dd", "E:\SCRIPTS\RTS-Dd\RTS-Dd", "C:\Windows\System32\WindowsPowerShell\v1.0\Modules\RTS-Dd\RTS-Dd"),
                    ("Asset MGMT", "E:\SCRIPTS\RTS-ASSET\RTS-ASSET", "C:\Windows\System32\WindowsPowerShell\v1.0\Modules\RTS-ASSET\RTS-ASSET"),
                    ("Testing Function", "E:\SCRIPTS\RTS-TEST\RTS-TEST", "C:\Windows\System32\WindowsPowerShell\v1.0\Modules\RTS-TEST\RTS-TEST")
                )


$MAIL_CHARSET_PATTERN = '[^A-z0-9@._-]'

$RTS_OU = 'OU=RTS,OU=Units,DC=media,DC=int'
$RTS_USERS_OU = "OU=Users,$RTS_OU"
$INTERNAL_OU = "OU=Internal,$RTS_USERS_OU"
$NETTOYAGE_OU = "OU=1 Infomut_Expires,OU=13-Nettoyage,$($RTS_USERS_OU)"

enum EmployeeTypes {
    Internal = 1
    External = 3
}

Function Grep{
    [CmdletBinding()]
    Param(
        [Parameter(ValueFromPipeline=$false)][string]$query,
        [Parameter(ValueFromPipeline)]$data
     )     
    return $data | fl | Out-String -stream | Select-String $query
}

Function Is-ModuleImported($moduleName){
    return Get-Module | ? { $_.Name -eq $currModuleName }
}

#Source https://stackoverflow.com/questions/36650961/use-read-host-to-enter-multiple-lines
Function Read-HostMultiline([string]$prompt){
    Write-Prompt $prompt
    $result = while ($true){
        Read-Host | set r; if (!$r) {break}; $r
    }
    return $result
}

Function Import-O365Session([bool]$silent = $false) {
    if ($script:isO365Imported) {
        return
    }

    if ($null -eq (Get-Module ExchangeOnlineManagement -ListAvailable)) {
        #On check si le module "ExchangeOnlineManagement" est installé
        if ($silent -eq $false) {
            Write-Info "Module ExchangeOnlineManagement introuvable, installation..."
        }
        Install-ExchangeOnlineModule
    }

    if ([System.Net.ServicePointManager]::SecurityProtocol -ne [System.Net.SecurityProtocolType]::Tls12) {
        #Il faut set la version du protocole TLS à 1.2, sinon la connexion sera refusée
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
    }
    if ($silent -eq $false) {
        Write-Info "Connexion à Exchange Online en cours ..."
    }
    try {
        #--------------------- L'ouverture de sessions Office 365 via New-PSSession est deprecated et ne fonctionnera pas ---------------------
        Connect-ExchangeOnline -UserPrincipalName ($env:USERNAME + "@rts.ch") -ShowBanner:$false
        $script:isO365Imported = $true
        return
    }
    Catch {
        Write-Host "Erreur lors de la connexion à Exchange Online, fermeture..."
        throw
    }
}
Function Get-CASSession([bool]$silent = $false, [bool]$forceImport = $false) {
    if ($null -ne $script:CASSession -and $script:CASSession.state -eq 'Opened') {
        return
    }

    if ($silent -eq $false) {
        Write-Info "Connexion au CAS..."
    }

    try {
        $script:CASSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://cas.media.int/PowerShell/ -Credential (get-credential $env:USERNAME)
        return 
    }
    Catch {
        Write-Err "Erreur lors de la connexion au CAS, fermeture..."
        exit
    }
}
Function Send-CASRequest($scriptBlock, [array]$argumentList) {
    Get-CASSession -silent:$true
    return Invoke-Command -Session $script:CASSession -ScriptBlock $scriptBlock -ArgumentList $argumentList
}

Function Get-CNFromDistinguishedName($distName){
    ($distName -match 'CN=(.+?),') | Out-Null #Out-Null est nécessaire, sinon la valeur de -match est aussi retournée (???)
    if($Matches.Count -gt 0){
        return $Matches[1]
    }
}

Function Ask-Confirmation([string]$prompt) {
    process {
        $prompt = Assert-StrArg $prompt $_ ''
        do {
            Write-Prompt "$prompt`n[O] Oui [N] Non: " 
            $answer = (Read-Host).ToLower().Trim()
        } While ($answer[0] -ne 'o' -and $answer[0] -ne 'n')
        return $answer[0] -eq 'o'
    }
}
Function Ask-Redo([string[]]$choices, $prompt = "Recommencer ?") {
    if (Ask-Confirmation $prompt) {
        if($choices.Length -eq 0){
            return $true
        }
            $choices += 'Aucun'
    $letters = $choices | % { $_.ToLower()[0] }
    $choices[0..($choices.Length - 2)] | % { $txt = "Garder la même" } { $txt += " $(Format-FirstChar -string $_)," }
    $txt = "$($txt.Substring(0,$txt.Length - 1)) ou $(Format-FirstChar -string $choices[$choices.Length-1])"
        
    do {
        Write-prompt "$($txt): "-NoNewLine:$true
        $val = (Read-Host).ToLower()[0]
    } while ($letters.IndexOf($val) -eq -1)

    Write-Host #On ajoute une ligne d'espacement
    return $val
    }
    else {
        return $false
    }
}
Function Request-MultipleChoices([string]$header, [string[]]$choices, [bool]$numberChoices = $false, [bool]$allowAll = $false, $newLineBefore = $false, [bool]$sortEntries = $false){
    $shortcutLength = 0

    if($sortEntries){
        $choices = $choices | Sort-Object
    }

    if($allowAll){
        $choices += $REQUEST_ALL #+= est le seul moyen d'ajouter des éléments à des arrays (taille fixe contrairement aux listes)
    }
    if($numberChoices){
        $shortcutLength = $choices.Count.ToString().Length
        $choices = $choices | % {$i = 0} { $i += 1; "$($i.ToString("D$shortcutLength")) $_" } #On ajoute des 0 au début si le chiffre n'est pas aussi long que les autres (Eg. '50'.ToString(D3) -> 050)
        #Si tout les choix ne sont pas de la même taille, il peut y avoir des ambiguités au niveau des choix. De plus, Format-FirstChar a été conçue pour fonctionner avec une outline de taille fixe
        $letters = (1..$choices.Count) | % { $_.ToString("D$shortcutLength") } #On génère un tableau contenant les nombres de 1 au nombre de choix, puis on converti en string
    }else{
        $duplicates = [System.Collections.ArrayList]@()
        $biggestChoice = ($choices | Sort-Object -Property { $_.Length } -Descending)[0].Length
        do{
            if($shortcutLength -gt $biggestChoice){
                Write-Err "Erreur lors de l'attribution de shortcuts..."
                return
            }
            $letters = $choices | % { $_.ToLower()[0..$shortcutLength] -join '' }
            $duplicates = $letters | Group-Object | ? { $_.Count -gt 1 } #Créé un tableau duplicates contenant toutes les lettres dupliquées
            $shortcutLength += 1
        }while($duplicates.Count -gt 0)
    }
    if($newLineBefore){
        Write-Host '`n' -NoNewline:$true
    }
    Write-Prompt "$($header.ToUpper())`n" #On mets une newline avant, afin que le prompt se démarque mieux
    $choices | % { Write-Prompt (Format-FirstChar "$_`n" -nbChars $shortcutLength) }

    do {
        Write-Prompt "Choix: " -noNewLine:$true
        $val = (Read-Host).ToLower()
    } while ($letters.IndexOf($val) -eq -1)
    if($allowAll -and $letters.IndexOf($val) -eq ($choices.Length - 1)){ #"Tous" sera toujours le dernier choix
        return $REQUEST_ALL
    }else{
        return $letters.IndexOf($val)
    }
}

Function Assert-StrArg {
    Param([Parameter(Mandatory = $true)][AllowNull()][AllowEmptyString()][string]$text, 
        [Parameter(Mandatory = $true)][AllowNull()][AllowEmptyString()][string]$pipeVal,
        [AllowNull()][AllowEmptyString()][string]$prompt = '')
    #Les paramètres sont nécessaires, sinon powershell va skipper les arguments vides
    #Permet de s'assurer qu'un argument n'est pas vide. Si $text et $pipeVal sont vides, l'utilisateur sera prompté pour la valeur
    #$text -> L'argument
    #$pipeVal -> La valeur potentielle dans le pipe (au cas-ou l'argument aurait été pipé plutôt que passé explicitement)
    #$prompt -> Le texte à afficher à l'utilisateur pour lui demander la valeur, si celle-ci n'a pas été pipée ou passée
    if ($text.Length -le 0) {
        if ($pipeVal.Length -gt 0) {
            return $pipeVal
        }
        else {
            if ($prompt.Length -gt 0) {
                Write-Prompt ($prompt + ": ")
                return (Read-Host)
            }
            else {
                return ''
            }
        }
    }
    else {
        return $text
    }
}

Function Prompt-ForDate([string]$prompt, [string]$regex = '[0-3][0-9].[0-1][0-9].(?:20)?[0-9]{2}', [bool]$canBeNull = $false) {
    do {
        Write-Prompt ($prompt + ": ")
        $date = Read-Host
        if($canBeNull -and $date.Length -eq 0){
            return $date
        }
    }While ((Test-DateFormat -date $date -regex $regex) -eq $false)
    return [DateTime]::Parse($date)
}

Function Remove-Accents([string]$text){
        $nom = Assert-StrArg $text $_
        If ($nom.Contains("ä")) 
        { $nom = $nom.Replace("ä", "ae") }	
        If ($nom.Contains("é"))
        { $nom = $nom.Replace("é", "e") }	
        If ($nom.Contains("ë"))
        { $nom = $nom.Replace("ë", "e") }	
        If ($nom.Contains("ê"))
        { $nom = $nom.Replace("ê", "e") }	
        If ($nom.Contains("ï"))
        { $nom = $nom.Replace("ï", "i") }	
        If ($nom.Contains("ü"))
        { $nom = $nom.Replace("ü", "ue") }	
        If ($nom.Contains("ù"))
        { $nom = $nom.Replace("ù", "u") }
        If ($nom.Contains("è"))
        { $nom = $nom.Replace("è", "e") }	
        If ($nom.Contains("ö"))
        { $nom = $nom.Replace("ö", "oe") }	
        If ($nom.Contains("ô"))
        { $nom = $nom.Replace("ô", "o") }
        If ($nom.Contains(" "))
        { $nom = $nom.Replace(" ", "") }
        return $nom
}

Function Format-Name([string]$nom) {
    process {
        $nom = Remove-Accents (Assert-StrArg $nom $_ "Veuillez indiquer le nom")
        return $nom
    }
}

Function Format-Mail([string]$mail, [bool]$stripDomain = $true){
process{
        $mail = Assert-StrArg $mail $_
        if($stripDomain -and $mail.Contains('@')){
            $mail = $mail.Split('@')[0]
        }
        return $mail -replace $MAIL_CHARSET_PATTERN
    }
}

Function Format-SAM([string]$nom, [string]$prenom, [bool]$isUser = $true) {
    if($isUser){
        $nom = Assert-StrArg $nom $_
        $nom = Remove-Accents ($nom.Replace('-',''))

        $prenom = Assert-StrArg $prenom $_
        $prenom = Remove-Accents ($prenom.Replace('-',''))
        if ($nom.Length -le 6 -or $nom.Length -eq 6) {
            return $nom + $prenom.Substring(0, 2)
        }
        else {
            return $nom.Substring(0, 6) + $prenom.Substring(0, 2) 
        }
    }else{
        $nom = Remove-Accents $nom
        $prenom = Remove-Accents $prenom
        $fullName = $nom + $prenom
        if($fullName.Length -gt 20){
            return $fullName.Substring(0,20)
        }else{
            return $fullName
        }

    }
}

Function Format-FirstChar([string]$string, $outline = "[]", [int]$nbChars = 1) {
    return $outline[0] + ($string[0..($nbChars - 1)] -join '').ToUpper() + $outline[1] + $string.Substring($nbChars)
}
Function Read-ForDate([string]$prompt, [string]$regex = '[0-3][0-9].[0-1][0-9].(?:20)?[0-9]{2}', [bool]$canBeNull = $false) {
    do {
        Write-Prompt ($prompt + ": ")
        $date = Read-Host
        if($canBeNull -and $date.Length -eq 0){
            return $date
        }
    }While ((Test-DateFormat -date $date -regex $regex) -eq $false)
    return $date
}
Function Wait-ProgressBar([int]$time, [string]$text = "En attente de synchronisation") {
    $loopNumber = 0
    if ($time -le 1) {
        Start-Sleep $time
        return
    }
    Write-Progress -Activity $text -Status "$time secondes restantes" -PercentComplete 0 #On initialise la barre de chargement
    while ($loopNumber -le $time) {
        Start-Sleep 1
        $loopNumber += 1
        $percentage = $loopNumber / $time * 100
        if ($percentage -gt 100) {
            #En raison de la façon qu'à powershell de gérer les divisions de nombres à virgule, $percentage peut parfois être > 100
            $percentage = 100
        }
        Write-Progress -Activity $text -Status "$($time - $loopNumber) secondes restantes" -PercentComplete $percentage
    }
    Write-Progress -Activity $text -Status "OK" -PercentComplete 100
    Write-Progress -Activity $text -Status "OK" -Completed
    return
}
Function Test-DateFormat([string]$date, [string]$regex = '[0-3][0-9].[0-1][0-9].(?:20)?[0-9]{2}') {
    return $date -match $regex
}
Function Underline([string]$txt) {
    return "$([char]27)[4m" + $txt + "$([char]27)[0m" #Souligne le texte via les séquences d'échappement ANSI
}

Function Write-Err([string]$text, [bool]$noNewLine = $false, [bool]$extraNewLine = $true) {
    $text = Assert-StrArg $text $_
    Write-Host $text -f Red -BackgroundColor Black -NoNewline:$noNewLine
    if($extraNewLine){
        Write-Host #On ajoute un retour à la ligne
    }
} 
Function Write-Info([string]$text, [bool]$noNewLine = $false) {
    $text = Assert-StrArg $text $_
    Write-Host $text -f Cyan -NoNewline:$noNewLine
}
Function Write-Success([string]$text, [bool]$noNewLine = $false) {
    $text = Assert-StrArg $text $_
    Write-Host $text -f Green -BackgroundColor Black -NoNewline:$noNewLine
}
Function Write-Prompt([string]$text, [bool]$noNewLine = $true) {
    Write-Host $text -f $PROMPT_COLOR -NoNewline:$noNewLine
}

# Module helpers

Function Update-RTSModules($module = $null, [bool]$recursiveCall = $false) {
    Begin {
        $now = (Get-Date -Format  "yyyy.MM.dd_HH.mm.ss")
        $ALL_MODULES = $ALL_MODULES | Sort-Object
    }
    Process {
        if($null -eq $module){
                $moduleNames = $ALL_MODULES | % { $_[0] }
                $moduleIndex = (Request-MultipleChoices "which module to update ?" $moduleNames -numberChoices:$true -allowAll:$true -sortEntries:$false)
                if($moduleIndex -eq $REQUEST_ALL){
                    return $moduleNames | % { Update-RTSModules $_ -recursiveCall:$true }
                }else{
                    $module = $ALL_MODULES[$moduleIndex]
                }
            }elseif ($module.GetType() -eq [string]){
                $module = ($ALL_MODULES | ? { $_[0] -like "*$module*" })
            }

        Copy-Item "$($module[2]).psm1" -Destination "$($module[2]).$now.psm1.bak"
        Copy-Item "$($module[1]).ps1" -Destination "$($module[2]).psm1" -Force
        Write-Success "Module $($module[0]) mis à jour"
    }
    end {
        if($recursiveCall -eq $false -and $moduleIndex -ne $REQUEST_ALL -and (Ask-Confirmation "Recommencer ?")){
            Update-RTSModules
        }
    }
}

Function Update-BossyModules {
    $path = ("D:\Bossyst\Script\AD\prod\CopyGroups", "C:\Windows\System32\WindowsPowerShell\v1.0\Modules\RTS-Bossy\RTS-Bossy")
    #Copie le script actuel dans le dossier modules de powershell, afin que les fonctions soient disponibles dans le 'path' powershell
    #A lancer lorsque le script a été modifié
    $formattedDate = Get-Date -Format "yyyy.MM.dd_HH.mm.ss"
    Copy-Item -Path ($path[1] + '.psm1') -Destination "$($path[1]).$formattedDate.psm1.bak" #On créé un backup du fichier
    Copy-Item -Path "$($path[0]).ps1" -Destination "$($path[1]).psm1" -Force
}

Function Update-DdModules {
    $path = ("D:\DavidD\Modules\RTS-Dd", "C:\Windows\System32\WindowsPowerShell\v1.0\Modules\RTS-Dd\RTS-Dd")
    #Copie le script actuel dans le dossier modules de powershell, afin que les fonctions soient disponibles dans le 'path' powershell
    #A lancer lorsque le script a été modifié
    $formattedDate = Get-Date -Format "yyyy.MM.dd_HH.mm.ss"
    Copy-Item -Path ($path[1] + '.psm1') -Destination "$($path[1]).$formattedDate.psm1.bak" #On créé un backup du fichier
    Copy-Item -Path "$($path[0]).ps1" -Destination "$($path[1]).psm1" -Force
}


#Fonctions RTS

Function Get-ADUserByUsername([string]$username, [boolean]$promptUserIfNotFound = $true) {
    process {
        $username = Assert-StrArg $username $_ ''
        $identity = $(Get-ADUser -filter "samaccountname -eq '$username'" -SearchBase $RTS_USERS_OU)

        if ($promptUserIfNotFound -eq $false) {
            return $identity
        }

        While ($null -eq $identity) {
            Write-Err "User $username introuvable"
            Write-Prompt "Username: "
            $nom = Read-Host
            $identity = $(Get-ADUser -Filter "SamAccountName -eq '$nom'" -SearchBase $RTS_USERS_OU)
        }
        return $identity
    }
}
Function Get-ADUserByMail([string]$mail, [boolean]$promptUserIfNotFound = $true) {
    process {
        $mail = Assert-StrArg $mail $_ ''
        $identity = $(Get-ADUser -filter "Mail -eq '$mail'" -SearchBase $RTS_USERS_OU)

        if ($promptUserIfNotFound -eq $false) {
            return $identity
        }

        While ($null -eq $identity) {
            Write-Err "User $nom introuvable"
            Write-Prompt "Mail: "
            $mail = Read-Host
            $identity = $(Get-ADUser -Filter "SamAccountName -eq '$mail'" -SearchBase $RTS_USERS_OU)
        }
        return $identity
    }
}

Function Wait-UntilMailboxSynced([string]$emailAddress, [int]$retryTimeout = 30, [bool]$silent = $false){ #Attends qu'une boîte mail sois trouvable sur O365
        $counter = 0
        Import-O365Session -silent:$true
        if($silent -eq $false){
            Write-Info "`nEn attente de synchronisation avec Office 365, cela peut prendre jusqu'à 60 minutes..."
        }
        $mailbox = Get-Mailbox -Filter "WindowsEmailAddress -eq '$emailAddress'"

        while($null -eq $mailbox){
            Wait-ProgressBar -time $retryTimeout -text "En attente... ($($retryTimeout)s)"

            $counter += 1
            $timeSpent = $retryTimeout / 60.0 * $counter
            $minutes = [Math]::Floor($timeSpent)
            $seconds = ($timeSpent % 1) * 60.0 #On récupère la partie fractionnelle ($timeSpent % 1) puis on le convertit en l'équivalent en minutes (Ex. 0.5 minutes = 30s)
            if($silent -eq $false){
                Write-Info "$($minutes.ToString("00")):$($seconds.ToString("00")) " -noNewLine:$true #"D" convertit un int en string, "2" indique que l'on veut un chiffre avec deux nombres (Ex. 5 -> 05)
            }

            $mailbox = Get-Mailbox -Filter "WindowsEmailAddress -eq '$emailAddress'"
        }
        if($silent -eq $false){
            Write-Host `n -NoNewline
        }
}

Function Install-ExchangeOnlineModule() {
    $ProgressPreference = "SilentlyContinue" #Pour ne pas prompter l'utilisateur
    try {
        Install-Module -Name ExchangeOnlineManagement -Force -Confirm:$false -ErrorAction Stop
    }
    catch {
        Write-Host "Erreur lors du téléchargement du module, merci de vous connecter à un réseau externe et de désactiver le proxy, puis réessayez..."
        throw-Error "Intall-Exchange-Online-Module : Téléchargement bloqué, le PC est probablement connecté au réseau interne"
    }
    $ProgressPreference = "Continue" #On remet les préfèrences comme elles l'étaient
}

#Source https://den.dev/blog/powershell-windows-notification/
Function Show-Notification {
    [cmdletbinding()]
    Param (
        [string]
        $ToastTitle,
        [string]
        [parameter(ValueFromPipeline)]
        $ToastText
    )

    [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
    $Template = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::ToastText02)

    $RawXml = [xml] $Template.GetXml()
    ($RawXml.toast.visual.binding.text | where { $_.id -eq "1" }).AppendChild($RawXml.CreateTextNode($ToastTitle)) | Out-Null
    ($RawXml.toast.visual.binding.text | where { $_.id -eq "2" }).AppendChild($RawXml.CreateTextNode($ToastText)) | Out-Null

    $SerializedXml = New-Object Windows.Data.Xml.Dom.XmlDocument
    $SerializedXml.LoadXml($RawXml.OuterXml)

    $Toast = [Windows.UI.Notifications.ToastNotification]::new($SerializedXml)
    $Toast.Tag = "PowerShell"
    $Toast.Group = "PowerShell"
    $Toast.ExpirationTime = [DateTimeOffset]::Now.AddSeconds(30)

    $Notifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("PowerShell")
    $Notifier.Show($Toast);
}

Function Format-StringFromTemplate($template, $values){
    $result = $template
    if($values[0] -is [String]){
        return $result.Replace($values[0], $values[1])
    }else{
        foreach($value in $values){
            if(Search-StringInArray $result $value[0]){
                $result = Format-StringFromTemplate $result $value
            }
        }
        return $result
    }
}

function Send-MailFromTemplate([string]$subject, $from, $to, [string]$header, [System.Collections.ICollection]$tableData) {
    $TOKEN_LABEL = "{LABEL}"
    $TOKEN_VALUE = "{VALUE}"

    $HEADER_TEMPLATE = @"
        <table style="width: 60%" style="border-collapse: collapse; border: 1px solid #8d0707;">

        <thead>
            <tr>
                <th colspan="2" bgcolor="#008080" style="color: #FFFFFF; font-size: large; height: 35px;">
                $TOKEN_LABEL</th>
            </tr>
        </thead>    
        <tbody>
"@

    $CELL_TEMPLATE =         
    @"
    <tr style="border-bottom-style: solid; border-bottom-width: 1px; padding-bottom: 1px">
        <td style="width: 201px; height: 35px">$TOKEN_LABEL</td>
        <td style="text-align: center; height: 35px; width: 233px;">
        <b>$TOKEN_VALUE</b></td>
    </tr>   
"@
    $body = ($HEADER_TEMPLATE.Replace($TOKEN_LABEL, $header)) +
            ($tableData | ? {$_ -ne $null -and $_.Length -eq 2} | % { $CELL_TEMPLATE.Replace($TOKEN_LABEL, $_[0]).Replace($TOKEN_VALUE, $_[1]) }) -join "`n" +
            ("</tbody></table>")

    Send-MailMessage -From $from -to $to -Subject $subject -Body $body -BodyAsHTML -encoding "Default" -SmtpServer cas.media.int
}

Function Make-DictionaryFromArray($array){
    $result = [System.Collections.ArrayList]@()
    for($i = 0; $i -lt $array.Count;$i += 2){
        $result.Add(@($array[$i], $array[$i+1])) | Out-Null
    }
    return $result
}

Function Search-StringInArray($array, [string]$string){
    $result = $false
    foreach($val in $array){
        if($val -is [System.Collections.ICollection]){
            $result = ($result -or (Search-StringInArray $val $string))
        }elseif($val -is [String]){
            $result = ($result -or ($val.Contains($string)) -or ($string.Contains($val)))
        }
    }
    return $result
    
}

Function Create-InfomutLogFile($user, [bool]$silent = $false){
    $line = Make-InfomutLogString $user
    if($silent -eq $false){
        Write-Success $line
        "$($env:USERNAME) - $line" >> "C:\Infomut\$($user.ADUser.SamAccountName).log" #'>' redirige l'output dans le fichier indiqué; '>>' redirige l'output, et l'ajoute à la fin du fichier spécifié
    }
}

Function Make-InfomutLogString($user){
    return "[$(Get-Date -Format 'yyyy.MM.dd hh.mm')] - [$($user.ticketNumber)] - [$($user.ADUser.UserPrincipalName)] Set expiration date $($user.expDate)"
}

$GRACE_PERIOD = 0
Function Process-InfomutUser($user){
    Set-ADAccountExpiration $user.ADUser.DistinguishedName ($user.expDate.AddDays($GRACE_PERIOD))
    $description = "$(Get-Date -format 'dd.MM.yy') - $($env:USERNAME) - Infomut $($user.ticketNumber)"
    if($null -ne $user.ADUser.Description){
        $description = "$description; $($user.ADUser.Description)"
    }
    Set-ADUser $user.ADUser.DistinguishedName -Description $description -Replace @{extensionAttribute11 = (Get-Date -Format "dd/MM/yy").ToString()}
}

Function Make-LeaversFromCSV([string]$filePath, [string]$delimiter = [System.Globalization.CultureInfo]::CurrentCulture.TextInfo.ListSeparator){
    $rawFile = Import-Csv -LiteralPath $filePath -Delimiter $delimiter #| % { [Leaver]::new((Get-ADUser $_.samAccountName), [DateTime]::Parse($_.expDate)) }
    $result = [System.Collections.ArrayList]@()

    foreach($entry in $rawFile){
        $name = ($entry.'Bénéficiaire'.Split(',')) | % { $_.Trim() }
        $mail = "$(Remove-Accents $name[1]).$(Remove-Accents $name[0])@rts.ch"
        $ticketNum = $entry.'N°'

        $user = Get-ADUser -Filter "UserPrincipalName -like '$mail'" -Properties ExtensionAttribute9, ExtensionAttribute11, Description -SearchBase $RTS_USERS_OU #On essaye d'abord de récupérer par le mail (Query rapide)

        if($null -eq $user){
            $user = Get-ADUser -Filter "Surname -like '*$($name[0])*' -and GivenName -like '*$($name[1])*'" -SearchBase $RTS_USERS_OU -Properties ExtensionAttribute9, ExtensionAttribute11, Description #Si on ne trouve pas, on essaye via lastName + firstName (Lent)
                if($null -eq $user){
                    Write-Err "[$ticketNum] Pas réussi à trouver l'utilisateur $($name[0]), $($name[1])" -noNewLine:$true
                    continue
                }
        }
        if($user.ExtensionAttribute11 -ne $null){
            Write-Info "User $($user.SamAccoutName) déjà traité le $($user.ExtensionAttribute11)"
            continue #Si l'attribut 11 est déjà set, c'est que l'utilisateur a déjà été traité un autre jour
        }
        $date = [DateTime]::MinValue
        try{
            if([DateTime]::TryParse($entry.'Date supplémentaire', [ref]$date) -or [DateTime]::TryParseExact($entry.'Date supplémentaire', 'dd.MM.yyyy', [ref]$date)){
                $result.Add([Leaver]::new($user, $date.AddDays($EXTRA_DAYS), $entry.'N°')) | Out-Null
                Write-Success "[$ticketNum] Parsé $($user.UserPrincipalName)"
            }else{
                Write-Err "[$ticketNum] Pas réussi à parser la date pour $($user.UserPrincipalName)" -noNewLine:$true
                continue
            }
        }catch{
            Write-Err "[$ticketNum] Pas réussi à parser la date pour $($user.UserPrincipalName)" -noNewLine:$true
            continue
        }
    }
    return $result
}

Function Get-FilePathFromDialog(){
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
    $dlg = [System.Windows.Forms.OpenFileDialog]::new()
    $dlg.Filter = "CSV (*.csv)|*.csv|All files (*.*)|*.*"
    $dlg.RestoreDirectory = $true
    if($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
        return $dlg.FileName
    }else{
        return $null
    }
}

Function Remove-DateFromCSV{
    $versions = @{
        '10' = 'Catalina'
        '11' = 'Big sur'
        '12' = 'Monterey'
        '13' = 'Ventura'
        '14' = 'Sonoma'
        '15' = 'Sequoia'
    }
    Write-Host "Séléctionner l'export EZV" -noNewLine:$false
    $file = Get-FilePathFromDialog
    if($null -eq $file){
        Write-Host "Veuillez séléctionner un fichier" -noNewLine:$true
        return    
    }
    
    $data = Import-Csv $file -UseCulture
    $data = $data | % {
        if($_.'OS version' -match '(.+\..+)\.20..'){
            $_.'OS version' = $Matches[1]
        }
        $_.OS = $versions[$_.'OS version'.substring(0,2)]
        return $_
    }
    
    $data | Select * | Export-Csv -Encoding UTF8 -NoTypeInformation -UseCulture -Path "$file.export.csv"
}

Function Get-ExpiredInfomutUsers([int]$gracePeriod){
    #Si la date d'expiration a été set il y a DAYS_BEFORE_NETTOYAGE jours, et que le compte est expiré, on le retourne
    $cleanupLimit = (Get-Date).AddDays(-$gracePeriod)
    $expiredUsers = (Search-ADAccount -AccountExpired -SearchBase $RTS_OU |`
     % { Get-ADUser $_.SamAccountName -Properties AccountExpirationDate, ExtensionAttribute11, ExtensionAttribute12 } |`
     ? { $_.ExtensionAttribute11.Length -gt 0 -and $_.ExtensionAttribute12.Length -eq 0 })
    return ($expiredUsers | ? { [DateTime]::Parse($_.ExtensionAttribute11) -le $cleanupLimit -and $_.AccountExpirationDate -le $cleanupLimit})
}

Function Get-LeaverAccounts ([int]$gracePeriod) {
    $leaverLimit = (Get-Date).AddDays(-$gracePeriod)
    return Get-ADUser -Filter '*' -SearchBase $NETTOYAGE_OU -Properties ExtensionAttribute12 | ? { $_.ExtensionAttribute12.Length -gt 0 -and [DateTime]::Parse($_.ExtensionAttribute12) -le $leaverLimit }
}

Function Remove-AllGroups([string]$sam){
    Get-ADUser $sam -Properties MemberOf | Select -ExpandProperty MemberOf | % { $group = $_ ` #On créé une variable "group", car sinon la valeur du pipe ($_) n'existera plus dans le catch
        try { Remove-ADGroupMember -Identity $_ -Members $sam -Confirm:$false -ErrorAction Stop }
        catch {Write-Err "[$(Get-CNFromDistinguishedName $group)] Error : $_" -noNewLine:$true} 
    }
}

Function Clear-UserPersonalData([string]$sam){
    $sam = Assert-StrArg $sam $_ "Entrez le nom d'utilisateur"
    Set-ADUser -Identity $sam -Clear EmployeeID, ExtensionAttribute10, ExtensionAttribute14, TelephoneNumber, Homephone, Mobile, IpPhone
}

Function Get-OUFromUserDistinguishedName([string]$dn){
    return $dn.Substring($dn.IndexOf('OU'))
}

Function Is-UserDistinguishedName([string]$val){
    return ($val.Contains("CN=") -and $val.Contains("OU=") -and $val.Contains("DC="))
}

Function New-RandomPassword([int]$length, [string]$charSet = "abcdefghjkmnpqrstuvwxyzABCDEFGHIJKLMNPQRSTUVWXYZ._-@+*%&?123456789"){
    $result = (1..$length) | % { $charSet[(Get-Random $charSet.Length)] }    
    return -join $result
    <#
    $charSet = $charSet.ToCharArray()
    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    $bytes = New-Object ($length)
    $rng.GetBytes($bytes)
    $result = New-Object char[]($length)
  
    for ($i = 0 ; $i -lt $length ; $i++) {
        $result[$i] = $charSet[$bytes[$i]%$charSet.Length]
    }
    return -join $result
    #>
}

Function Check-PasswordConformity([string]$pwd) {
    if ($pwd.Length -le 11) {
        return $false
    }

    $hasLowercase = $pwd -match '[a-z]'
    $hasUppercase = $pwd -match '[A-Z]'
    $hasDigits = $pwd -match '[0-9]'
    $hasSpecialChar = $pwd -match '[^A-z0-9]'

    return $hasLowercase -and $hasUppercase -and ($hasDigits -or $hasSpecialChar)
}


Function Generate-RandomPassword($length = 12) {
    #Génère un mot de passe qui contient toujours une majuscule, un chiffre et un caractère spécial
    $dictRegular = "abcdefghijklmnopqrstuvwxyz"
    $dictCapital = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    $dictNumbers = "0123456789"
    $dictSpecial = "._-@+*%&?!"
    $r = New-Object System.Random

    $result = (0..($length - 1)) | % { return $dictRegular[$r.Next(0, $dictRegular.Length)] }
    $result[0] = $dictCapital[$r.Next(0, $dictCapital.Length)]

    $result[$length - 2] = $dictNumbers[$r.Next(0, $dictNumbers.Length)]
    $result[$length - 1] = $dictSpecial[$r.Next(0, $dictSpecial.Length)]
    return $result -join ""
    
}

Function Get-PermutationsForPhoneNumber([string]$number){
    
    #On strip tout les charactères non numériques (et on garde le + du début si celui-ci est présent)
    $number = $number -replace '[^\+\d]'

    #Si le num ne commence ni par + ni par 0 c'est qu'il est au format d'openscape (par exemple 41581342323), on ajoute un + pour avoir le format international
    if($number[0] -ne '+' -and $number[0] -ne '0'){
        $number = '+' + $number
    }

    #Si le numéro commence par 00 on remplace par un +, pour n'avoir qu'un format international au lieu de 2 (+ et 00)
    if($number.StartsWith('00')){
        $number = '+' + $number.Substring(2)
    }

    #Si le numéro est Suisse mais qu'il est au format international, on transforme le format international en format national
    #Cela facilitera la reconstruction du numéro plus tard
    if($number.StartsWith('+41')){
        $number = '0' + $number.Substring(3)
    }

    $groups = ([Regex]::Match($number, $PHONE_REGEX)).Groups
    if($null -ne $groups -and $groups.Count -gt 0){
        #On reconstruit le numéro avec différents formats : Exemple 0788542432
        #1. 078 854 24 32
        #2. 0788542432
        #3. +41 78 854 24 32
        #4. 0041 78 854 24 32
        #5. +4178 854 24 32
        #6. +41788542432
        #7. 004178 854 24 32
        #8. 0041788542432
        $numbersToQuery = [System.Collections.ArrayList]::new()
        $ndc = $groups['NDC']
        $trio = $groups['Trio']
        $firstPair = $groups['FirstPair']
        $secondPair = $groups['SecondPair']
        $countryCode = $groups['International']
        if($groups['Local'].Success){
            #Formats nationaux
            $numbersToQuery.AddRange((
                [String]::Format("0{0} {1} {2} {3}", ($ndc, $trio, $firstPair, $secondPair)), #Format 1
                [String]::Format("0{0}{1}{2}{3}", ($ndc, $trio, $firstPair, $secondPair)) #Format 2
            ))
            $countryCode = '+41'
        }
        $numbersToQuery.AddRange((
            [String]::Format("{0} {1} {2} {3} {4}", ($countryCode, $ndc, $trio, $firstPair, $secondPair)), #Format 3
            [String]::Format("{0}{1} {2} {3} {4}", ($countryCode, $ndc, $trio, $firstPair, $secondPair)), #Format 5
            [String]::Format("{0}{1}{2}{3}{4}", ($countryCode, $ndc, $trio, $firstPair, $secondPair)), #Format 6
            [String]::Format("{0} {1} {2} {3} {4}", ($countryCode.Replace('+', '00'), $ndc, $trio, $firstPair, $secondPair)), #Format 4
            [String]::Format("{0}{1} {2} {3} {4}", ($countryCode.Replace('+', '00'), $ndc, $trio, $firstPair, $secondPair)), #Format 7
            [String]::Format("{0}{1}{2}{3}{4}", ($countryCode.Replace('+', '00'), $ndc, $trio, $firstPair, $secondPair)) #Format 8
        ))
    }
    return $numbersToQuery
}

Function Convert-ToGUID($uuid) {
    $uuid = Assert-StrArg $uuid $_ "Merci d'entrer l'UUID"
    return $uuid[6..7] + $uuid[4..5] + $uuid[2..3] + $uuid[0..1] + '-' + $uuid[10..11] + $uuid[8..9] + '-' + $uuid[14..15] + $uuid[12..13] + '-' + $uuid[16..19] + '-' + $uuid[20..31] -join ''
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
Function Get-PermutationsForPhoneNumber([string]$number) {
    $PHONE_REGEX = '^(?:(?<International>(?:\+?\d{2}))|(?<Local>0))(?<NDC>\d{2})(?<Trio>\d{3})(?<FirstPair>\d{2})(?<SecondPair>\d{2})'
    #On strip tout les charactères non numériques (et on garde le + du début si celui-ci est présent)
    $number = $number -replace '[^\+\d]'

    #Si le num ne commence ni par + ni par 0 c'est qu'il est au format d'openscape (par exemple 41581342323), on ajoute un + pour avoir le format international
    if ($number[0] -ne '+' -and $number[0] -ne '0') {
        $number = '+' + $number
    }

    #Si le numéro commence par 00 on remplace par un +, pour n'avoir qu'un format international au lieu de 2 (+ et 00)
    if ($number.StartsWith('00')) {
        $number = '+' + $number.Substring(2)
    }

    #Si le numéro est Suisse mais qu'il est au format international, on transforme le format international en format national
    #Cela facilitera la reconstruction du numéro plus tard
    if ($number.StartsWith('+41')) {
        $number = '0' + $number.Substring(3)
    }

    $groups = ([Regex]::Match($number, $PHONE_REGEX)).Groups
    if ($null -ne $groups -and $groups.Success) {
        #On reconstruit le numéro avec différents formats : Exemple 0788542432
        #1. 078 854 24 32
        #2. 0788542432
        #3. +41 78 854 24 32
        #4. 0041 78 854 24 32
        #5. +4178 854 24 32
        #6. +41788542432
        #7. 004178 854 24 32
        #8. 0041788542432
        $numbersToQuery = [System.Collections.ArrayList]::new()
        $ndc = $groups['NDC'].Value
        $trio = $groups['Trio'].Value
        $firstPair = $groups['FirstPair'].Value
        $secondPair = $groups['SecondPair'].Value
        $countryCode = $groups['International'].Value

        #if($null -ne $countryCode -and $countryCode[0] -ne '+' -and ){
        #    $countryCode = '+' + $countryCode
        #}
        if($groups['Local'].Success){
            $countryCode = '+41'
        }
                    
        if($null -eq $ndc -or $null -eq $trio -or $number -eq $firstPair -or $null -eq $secondPair -or $number -eq $countryCode){
            return $false
        }

        $numbersToQuery.AddRange((
                [String]::Format("{0} {1} {2} {3} {4}", ($countryCode, $ndc, $trio, $firstPair, $secondPair)), #Format 3
                [String]::Format("{0}{1} {2} {3} {4}", ($countryCode, $ndc, $trio, $firstPair, $secondPair)), #Format 5
                [String]::Format("{0}{1}{2}{3}{4}", ($countryCode, $ndc, $trio, $firstPair, $secondPair)), #Format 6
                [String]::Format("{0} {1} {2} {3} {4}", ($countryCode.Replace('+', '00'), $ndc, $trio, $firstPair, $secondPair)), #Format 4
                [String]::Format("{0}{1} {2} {3} {4}", ($countryCode.Replace('+', '00'), $ndc, $trio, $firstPair, $secondPair)), #Format 7
                [String]::Format("{0}{1}{2}{3}{4}", ($countryCode.Replace('+', '00'), $ndc, $trio, $firstPair, $secondPair)) #Format 8
            ))

        if ($groups['Local'].Success) {
            #Formats nationaux
            $numbersToQuery.AddRange((
                    [String]::Format("0{0} {1} {2} {3}", ($ndc, $trio, $firstPair, $secondPair)), #Format 1
                    [String]::Format("0{0}{1}{2}{3}", ($ndc, $trio, $firstPair, $secondPair)) #Format 2
                ))
        }

    }
    return $numbersToQuery
        
}
Using module RTS-Components #Nécessaire pour l'objet Checklist

<#
    Set-Place:
    Module:
    ExchangePowerShell
    Applies to:
    Exchange Online
    This cmdlet is available only in the cloud-based service.

    Use the Set-Place cmdlet to update room mailboxes with additional metadata, which provides a better search and room suggestion experience.

    Note: In hybrid environments, this cmdlet doesn't work on the following properties on synchronized room mailboxes:
    City, CountryOrRegion, GeoCoordinates, Phone, PostalCode, State, and Street. 
    To modify these properties except GeoCoordinates on synchronized room mailboxes, use the Set-User or Set-Mailbox cmdlets in on-premises Exchange.

    Tags exchange:
        Yealink (Optional BYOD, eg. A20, A30, etc...)
        Yealink (BYOD, eg. A10)
        Logitech
        Vidéo-projecteur
        MeetingBoard

        Cabine
        Salle

        WPP20
        WPP30

        Réservation Libre
        Validation nécessaire

        Accès libre
        Clef à la réception
#>

Function New-RTSConferenceRoom() {
    param (
        $RoomType = $null,
        [string]$location,
        [string]$officeNum,
        [string]$capacity,
        [string]$floor,
        [int]$reservationType = 0
    )
    Begin {
        do {
            if ($null -eq $RoomType) {
                $RoomTag = "Salle de conférence"
            }
            Write-Prompt "Emplacement ('GE', 'LA' ou 'CUSTOM'): "
            $location = (Read-Host).ToUpper() #Il est important d'appeler Assert-StrArg sans $location, car sinon le do{}while tournera en boucle
            switch ($location) {
                "GE" { $city = "Genève" }
                "LA" { $city = "Lausanne" }
                "CUSTOM" {
                    $location = ((Assert-StrArg '' '' "Lieu (2 lettres, eg. 'GE')")[0..2]).ToUpper()
                    $city = Assert-StrArg '' '' "Ville"
                }
            }
        }
        while ($null -eq $city -or $location.Length -ne 2)
        $officeNum = Assert-StrArg $officeNum '' "Numéro de bureau (Ex: 10L43 / 412)"
        $capacity = Assert-StrArg $capacity '' "Nombre de places"
        $floor = Assert-StrArg $floor '' "Numéro de l'étage"
        $roomName = "RTS-$location-$officeNum"
        $mail = "$roomName@rts.ch"

        $DISPLAY_NAME = "$roomName $RoomTag"
        $COMPANY = "Radio Télévision Suisse" 
        $MAILBOX_OU = "OU=10-Mailbox,OU=Users,OU=RTS,OU=Units,DC=media,DC=int"

        if ($reservationType -le 0 -or $reservationType -gt 3) {
            $reservationType = [int]::Parse((Request-MultipleChoices "Type de réservation: `n" @(
                        "1 AutoAccept : Tout le monde peut réserver cette salle"
                        "2 Approved   : Approbation par délègués"
                        "3 Restricted : Refus automatique sauf pour les users définis"
                    )))
            $reservationType += 1    
        }
        if ($reservationType -eq 1) {
            $reservationTag = "Réservation libre"
        }
        if ($reservationType -eq 2) {
            $delegate = (Read-HostMultiline "Username des délègués (1 par ligne): `n")
            $reservationTag = "Validation nécessaire"
        }
        if ($reservationType -eq 3) {
            $restrictedUser = (Read-HostMultiline "Délégués (liste de distribution ou username) (1 par ligne): `n")
            $reservationTag = "Accès restreint"
        }

        Write-Host "La salle est-elle équipée de matériel pour la visioconférence ?"
        $ConfirmHardware = Read-Host "[O]ui ou [N]on"

        if ($location -eq 'GE') {
            $group = "RTS-Rooms-Geneve"
            $address = "Quai Ernest-Ansermet 20"
            $zip = "1205"
            $state = "GE"
        }
        elseif ($location -eq 'LA') {
            $group = "RTS-Rooms-Lausanne"
            $address = "Avenue du Temple 40"
            $Zip = "1012"
            $state = "VD"
        }
        $checklist = [checklist]::new("Création de la ressource", $(
                "Ouverture de la session O365"
                "Connexion au CAS"
                "Création de la salle dans l'AD"
                "Réplication AD"
                "Synchronisation CAS (30s)"
                "Activation de la mailbox sur CAS"
                "Synchronisation O365 (5 à 60 minutes)"
                "Paramètrage de la mailbox"
                "Paramètrage du calendrier"
                "Paramètrage du type de réservation"
                "Ajout dans le groupe RTS-Rooms"
                "Ajout des tags"
            ))
    }
    Process {
    
        # START CHECKLIST
        $checklist.Start()

        # Ouverture de la session O365
        Import-O365Session -silent:$true
        $checklist.SetStateAndGoToNext($true)

        # Connexion au CAS
        Get-CASSession -silent:$true
        $checklist.SetStateAndGoToNext($true)

        # Création de la salle dans l'AD
        try {
            New-ADUser -Name $roomName -UserPrincipalName $mail -EmailAddress $mail -DisplayName $DISPLAY_NAME -Path $MAILBOX_OU -Company $COMPANY `
                -Office $officeNum -City $city -Description $RoomTag -StreetAddress $address -PostalCode $zip -State $state
            Set-ADUser -Identity $roomName -Add @{extensionAttribute1 = "$reservationTag" }
        }
        catch {
            $checklist.FatalError()
            Write-Err "Erreur FATAL lors de la création de l'utilisateur, fermeture..."
            Write-Err $_
            return
        }
        $checklist.SetStateAndGoToNext($true)

        # Réplication AD
        try {
            while ($null -eq (Get-ADUser -Filter "SamAccountName -eq '$roomName'")) {
                Start-Sleep 1
            }
        }
        catch {
            $checklist.FatalError()
            Write-Err "Erreur FATAL lors de la réplication AD"
            Write-Err $_
            return   
        }
        $checklist.SetStateAndGoToNext($true)

        # Synchronisation CAS
        try {
            Wait-ProgressBar -time 15 -text "En attente de synchro pour activation du lien exchange Onmicrosoft.com..."       
        }
        catch {
            $checklist.FatalError()
            Write-Err "Erreur FATAL lors de la Synchronisation CAS"
            Write-Err $_
            return   
        }
        $checklist.SetStateAndGoToNext($true)
 
        # TODO: Check for SRV "gds-ms-iaam043.media.int" to "gds-ms-iaam048.media.int" & Manage creation on each 1 == maybe foreach on SRV list
        # Activation de la mailbox sur CAS 
        try {
            Send-CASRequest -ScriptBlock { param($roomName, $mail) Enable-RemoteMailbox $mail -Room -DomainController "gds-ms-iaam043.media.int" -RemoteRoutingAddress "$roomName@SRGSSR.mail.onmicrosoft.com" } -argumentList $roomName, $mail | Out-Null
        }
        catch {
            $checklist.FatalError()
            Write-Err "Erreur FATAL lors de l'activation de la mailbox sur CAS"
            Write-Err $_
            return        
        }
        $checklist.SetStateAndGoToNext($true)

        # Synchronisation O365
        try {
            Wait-UntilMailboxSynced $mail -silent:$true
        }
        catch {
            $checklist.FatalError()
            Write-Err "Erreur FATAL lors du paramètrage de la synchronisation O365"
            Write-Err $_
            return        
        }
        $checklist.SetStateAndGoToNext($true)

        # Paramètrage de la mailbox
        try {
            Set-Mailbox -Identity $mail -AuditEnabled $true -AuditDelegate "Update, Move, MoveToDeletedItems, SoftDelete, HardDelete, SendAs, SendOnBehalf, Create" `
                -AuditOwner "Move, MoveToDeletedItems, SoftDelete, HardDelete" -RetentionPolicy "SRG SSR" -ResourceCapacity $capacity
            Set-MailboxMessageConfiguration -Identity $mail -IsReplyAllTheDefaultResponse $False
        }
        catch {
            $checklist.FatalError()
            Write-Err "Erreur FATAL lors du paramètrage de la mailbox"
            Write-Err $_
            return
        }
        $checklist.SetStateAndGoToNext($true)

        # Paramètrage du calendrier
        try {
            Set-MailboxCalendarConfiguration -Identity $mail -WeekStartDay Monday -WorkDays Weekdays -WorkingHoursStartTime 08:00:00 -WorkingHoursEndTime 17:00:00 `
                -WorkingHoursTimeZone "W. Europe Standard Time" -ShowWeekNumbers $true
            Set-Clutter -Identity $mail -Enable $false | Out-Null
            Set-CalendarProcessing -Identity $mail -AutomateProcessing "AutoAccept" -MaximumConflictInstances 20 -ConflictPercentageAllowed 35 -BookingWindowInDays 395 `
                -DeleteSubject $false -AddOrganizerToSubject $true -RemovePrivateProperty $false
        }
        catch {
            $checklist.FatalError()
            Write-Err "Erreur FATAL lors du paramètrage du calendrier"
            Write-Err $_
            return        
        }
        $checklist.SetStateAndGoToNext($true)

        # Paramètrage du type de réservation
        switch ($reservationType) {
            1 { Set-CalendarProcessing -Identity $mail -BookInPolicy @() -AllBookInPolicy $true } # AutoAccept for all users
            2 { Set-CalendarProcessing -Identity $mail -BookInPolicy @() -AllBookInPolicy $false -ResourceDelegates $delegate -AllRequestInPolicy $true -DeleteSubject $false -AddOrganizerToSubject $false } # Approved by delegates
            3 { Set-CalendarProcessing -Identity $mail -BookInPolicy @($restrictedUser) -AllBookInPolicy $false -DeleteSubject $false -AddOrganizerToSubject $false } # Restricted to BookInPolicy users 
            Default { Write-Err "Erreur vis-à-vis du type de réservation, fermeture..."; return }
        }
        $checklist.SetStateAndGoToNext($true)

        # Ajout dans le groupe RTS-Rooms 
        if ($null -ne $group) {
            Add-ADGroupMember -Identity $group -Members $roomName
        }
        $checklist.SetStateAndGoToNext($true)

        # Ajout des tags
        try {
            if($ConfirmHardware.ToUpper() -eq "N") {
                Set-Place -Identity $roomName -Floor $floor -Tags "$reservationTag","$RoomTag"
            }
            elseif($ConfirmHardware.ToUpper() -eq "O") {
                $hardwareTag = "OK"
                Set-Place -Identity $roomName -Floor $floor -AudioDeviceName "$hardwareTag" -VideoDeviceName "$hardwareTag" -DisplayDeviceName "$hardwareTag" -Tags "$reservationTag","$RoomTag"
            }
            else {
                Set-Place -Identity $roomName -Floor $floor -Tags "$reservationTag","$RoomTag" 
            }
        }
        catch {
            $checklist.FatalError()
            Write-Err "Erreur FATAL d'assignation des tags, cette pièce va s'autodetruire dans 5 secondes..."
            Write-Err $_
            continue
        }

        # END CHECKLIST
        $checklist.SetState($true) 
    }
    end {
        if (Ask-Confirmation "Recommencer ?") {
            New-RTSConferenceRoom
        }
    }
}


# TODO: Rework delegate
Function Add-BookingDelegate([string]$mailbox, [string]$user, [bool]$silent = $false) {
    Begin {
        $mailbox = Assert-StrArg $mailbox '' "Entrez l'alias ou l'email de la salle"
        if ($user.Length -le 0) {
            $user = Assert-StrArg $user $_ "Entrez le nom ou l'adresse mail de la personne"
            $user = (Get-ADUserByUsername $user).UserPrincipalName #On s'assure que la personne existe
        }

        Import-O365Session -silent:$true

        While ($null -eq $(Get-Mailbox -filter "userPrincipalName -eq '$mailbox' -or Alias -eq '$mailbox'")) {
            Write-Err  "`nLa boîte n'existe pas, merci de réessayer"
            Write-Prompt "Alias ou mail de la salle: "
            $mailbox = Read-Host
        }
        if (-not(Is-ModuleImported 'RTS-Components')) {
            Import-Module 'RTS-Components' | Out-Null
        }
        $checkList = [Checklist]::new("Ajout des droits de réservation", $(
                "Ajout des accès à la salle"
            ))
    }
    Process {
        #Set-CalendarProcessing -Identity $mailbox -ResourceDelegates $user.UserPrincipalName
        $checkList.Start()
        Add-UserToMailbox -user $user -mailbox $mailbox -silent:$true

        $alias = (Get-Mailbox $mailbox).Alias
        if ($null -ne (Get-ADGroup -Filter "SamAccountName -like 'RTS-DL-MAIL-$alias'")) {
            try {
                if ($user.Contains('@')) {
                    $user = (Get-ADUserByMail $user).SamAccountName
                }
                Add-ADGroupMember -Identity "RTS-DL-MAIL-$alias" -Members $user
            }
            catch {
                $checkList.SetState($false)
                Write-Err "Erreur lors de l'ajout à la liste 'RTS-DL-MAIL-$alias'"
                return
            }
        }

        $checklist.SetState($true)
    }
    End {
        while ($silent -eq $false -and $val -ne $false) {
            $val = Ask-Redo -choices ('boîte', 'personne')

            if ($val -ne $false) {
                if ($val -eq 'a') {
                    Add-BookingDelegate -silent $true
                }
                elseif ($val -eq 'b') {
                    Add-BookingDelegate -mailbox $mailbox -silent $true
                }
                else {
                    Add-BookingDelegate -user $user -silent $true
                }
            }
        }
    }
}

# TODO: Rework delegate
Function Add-MultipleBookingDelegate([string]$mailbox, [bool]$silent = $false) {
    Begin {

        if (-not(Is-ModuleImported 'RTS-Components')) {
            Import-Module 'RTS-Components' | Out-Null
        }
        
        $checkList = [Checklist]::new("Ajout des droits de réservation", $(
                "Ouverture de la session O365",
                "Choisir le .csv avec le samAccountName en Header",
                "Import des données",    
                "Ajouts multiple des accès à la salle"
            ))
    }
    Process {

        $mailbox = Assert-StrArg $mailbox '' "Entrez l'alias ou l'email de la salle"

        $checkList.Start() # Ouverture de la session O365
        Import-O365Session -silent:$true

        While ($null -eq $(Get-Mailbox -filter "userPrincipalName -eq '$mailbox' -or Alias -eq '$mailbox'")) {
            Write-Err  "`nLa boîte n'existe pas, merci de réessayer"
            Write-Prompt "Alias ou mail de la salle: "
            $mailbox = Read-Host
        }

        $alias = (Get-Mailbox $mailbox).Alias

        $checklist.SetStateAndGoToNext($true)# Choisir le .csv avec le samAccountName Header
        $fileName = Get-FilePathFromDialog

        try {
            $csvData = Import-Csv -Path $fileName -Delimiter ";" -Encoding "utf8"
        }
        catch {
            Write-Err "$_.Exception.Message"
            $checkList.SetState($false)
            return
        }

        $checklist.SetStateAndGoToNext($true)# Import des données

        foreach ($item in $csvData) {

            $userName = $item.samAccountName

            try {
                Add-UserToMailbox -user $userName -mailbox $mailbox -silent:$true 
            }
            catch {
                $checkList.SetState($false)
                Write-Err "$_.Exception.Message"
                return
            }
            
            if ($null -ne (Get-ADGroup -Filter "SamAccountName -like 'RTS-DL-MAIL-$alias'")) {

                try {
                    Add-ADGroupMember -Identity "RTS-DL-MAIL-$alias" -Members $userName
                }
                catch {
                    $checkList.SetState($false)
                    Write-Err "$_.Exception.Message"
                    Write-Err "Erreur lors de l'ajout à la liste 'RTS-DL-MAIL-$alias'"
                    return
                }
            }
        }
        $checklist.SetStateAndGoToNext($true) # Ajouts multiple des accès à la salle
        $checklist.SetState($true)
    }
    End {
        if (Ask-Confirmation "Recommencer ?") {
            Add-MultipleBookingDelegate
        }
    }
}

# TODO: Rework delegate
Function Remove-BookingDelegate([string]$mailbox, [string]$user, [bool]$silent = $false) {
    Begin {
        $mailbox = Assert-StrArg $mailbox '' "Entrez l'alias ou l'email de la salle"
        if ($user.Length -le 0) {
            $user = Assert-StrArg $user $_ "Entrez le nom ou l'adresse mail de la personne"
            $user = (Get-ADUserByUsername $user).UserPrincipalName #On s'assure que la personne existe
        }

        Import-O365Session -silent:$true

        While ($null -eq $(Get-Mailbox -filter "UserPrincipalName -eq '$mailbox' -or Alias -eq '$mailbox'")) {
            Write-Err  "`nLa boîte n'existe pas, merci de réessayer"
            Write-Prompt "Alias ou mail de la salle: "
            $mailbox = Read-Host
        }

        if (-not(Is-ModuleImported 'RTS-Components')) {
            Import-Module 'RTS-Components' | Out-Null
        }
        $checkList = [Checklist]::new("Retrait des droits de réservation", $(
                "Suppression des accès à la salle"
            ))
    }
    Process {
        
        $checkList.Start()
        Remove-UserFromMailbox -user $user -mailbox $mailbox -silent:$true #On retire les accès délégués

        $alias = (Get-Mailbox $mailbox).Alias
        if ($null -ne (Get-ADGroup -Filter "SamAccountName -like 'RTS-DL-MAIL-$alias'")) {
            #On retire aussi les accès spécifiques (Via liste de distribution) si ceux-ci existent
            try {
                if ($user.Contains('@')) {
                    #Add-ADGroupMember prend un SamAccountName et pas une addresse mail
                    $user = (Get-ADUserByMail $user).SamAccountName
                }
                Add-ADGroupMember -Identity "RTS-DL-MAIL-$alias" -Members $user
            }
            catch {
                $checkList.SetState($false)
                Write-Err "Erreur lors de l'ajout à la liste 'RTS-DL-MAIL-$alias'"
                return
            }
        }

        $checklist.SetState($true)
    }
    End {
        while ($silent -eq $false -and $val -ne $false) {
            $val = Ask-Redo -choices ('boîte', 'personne')

            if ($val -ne $false) {
                if ($val -eq 'a') {
                    Remove-BookingDelegate -silent $true
                }
                elseif ($val -eq 'b') {
                    Remove-BookingDelegate -mailbox $mailbox -silent $true
                }
                else {
                    Remove-BookingDelegate -user $user -silent $true
                }
            }
        }
    }
}

Function Set-ConferenceRoomAutoResponse([string]$mailbox, [bool]$enabled = $true, $response = $false) {
    begin {
        [checklist]$checklist = [checklist]::new("Réponse automatiques ConferenceRoom", (
                "Connexion O365",
                "Paramètrage des réponses automatiques"
            ))
        $mailbox = Assert-StrArg $mailbox $_ "Mailbox"
        if ($response.Length -eq $false) {
            $response = (Read-HostMultiline "Réponse automatique (multi-lignes):") -replace ('\r?\n', '<br>')
        }
    }     
    process {
        $checklist.Start()
        try {
            Import-O365Session -silent:$true
        }
        catch {
            $checklist.FatalError()
            Write-Err "Erreur lors de la connexion à O365"
            return
        }

        $checklist.SetStateAndGoToNext($true)
        try {
            if ($enabled -eq $false) {
                Set-CalendarProcessing -Identity $mailbox -AddAdditionalResponse $false
            }
            Set-CalendarProcessing -Identity $mailbox -AddAdditionalResponse $true -AdditionalResponse $response
        }
        catch {
            $checklist.FatalError()
            Write-Err "Erreur lors du paramètrage des réponses automatiques`n$($_.Exception)"
            return
        }
        $checklist.SetState($true)
        if (Ask-Confirmation "Recommencer ?") {
            Set-ConferenceRoomAutoResponse
        }
    }
}

# TODO:
# - Add Info-Tip with HTML & CSS
Function Enable-MTRMailbox ([string]$roomName, [bool]$silent = $false) {

    Begin {
        
        $DESCRIPTION = "Salle de conférence MTR /!\ NE PAS RETIRER LA LICENCE E3 DU COMPTE"
        $LICENCE_E3 = "RTS-G-L-M365_E3"
        $MTR_TAGS = "Salle de conférence MTR"

        while ($null -eq $room) {

            # Prompt for room name
            $roomName = Assert-StrArg $roomName '' "UserName de la salle"
            $room = Get-ADUser -identity $roomName
        }

        $sam = "$($room.SamAccountName)-MTR"
        $DISPLAY_NAME = "$($sam) Salle de conférence"
        $reservationType = (Get-ADUser -identity $roomName -Properties extensionattribute1 | Select-Object extensionAttribute1)
        $reservationType = $reservationType[0].extensionAttribute1

        # Create Password
        $PASSWORD = New-RandomPassword 20

        if (-not(Is-ModuleImported 'RTS-Components')) {
            Import-Module 'RTS-Components' | Out-Null
        }

        # Start O365 session
        Import-O365Session -silent:$true
    }
    Process {
        $checklist = [Checklist]::new("Activation de la salle de conférence MTR", $(
                "Renommage du compte",
                "Ajout du mot de passe ($PASSWORD)",
                "Activation du compte AD",
                "Ajout licence E3",
                "Ajout des tags"
            ))

        $checklist.Start() 

        # Renommage du compte
        try {
            Set-ADUser -Identity $room -UserPrincipalname "$sam@rts.ch"
            Rename-ADObject $room -NewName $DISPLAY_NAME
            # TODO: add UPN to aliases, pour l'instant il faut créer l'email alias à la main
        }
        catch {
            $checklist.FatalError()
            Write-Err "Erreur lors du renommage"
            Write-Err $_
            return
        }
        $checklist.SetStateAndGoToNext($true)


        # Ajout du mot de passe
        try {
            Set-ADAccountPassword -identity $roomName -Reset -NewPassword (ConvertTo-SecureString -AsPlainText "$PASSWORD" -Force) 
        }
        catch {
            $checklist.FatalError()
            Write-Err "Erreur lors de la création du mot de passe"
            Write-Err $_
            return
        }
        $checklist.SetStateAndGoToNext($true)


        # Activation du compte AD
        try {
            Set-ADUser -identity $room -Enabled:$True -Description $DESCRIPTION -DisplayName $DISPLAY_NAME -ChangePasswordAtLogon:$false -CannotChangePassword:$true -PasswordNeverExpire:$true
        }
        catch {
            $checklist.FatalError()
            Write-Err "Erreur lors de l'activation du compte AD"
            Write-Err $_
            return
        }
        $checklist.SetStateAndGoToNext($true)


        # Ajout licence E3
        try {
            Add-ADGroupMember -Identity $LICENCE_E3 -Members $room
        }
        catch {
            $checklist.FatalError()
            Write-Err "Erreur d'assignation du groupe"
            Write-Err $_
            return
        }
        $checklist.SetStateAndGoToNext($true)


        # Ajout des tags
        try {
            # TODO: Add ancient tags + $MTR_TAGS or Clear and reset ?
            $hardwareTag = "OK"
            Set-Place -Identity $roomName -AudioDeviceName "$hardwareTag" -VideoDeviceName "$hardwareTag" -DisplayDeviceName "$hardwareTag" -Tags "$reservationType","$MTR_TAGS"
        }
        catch {
            $checklist.FatalError()
            Write-Err "Erreur d'assignation des tags"
            Write-Err $_
            return
        }
        $checklist.SetState($true)

        Write-Info "Password du compte : $PASSWORD"
    }
    End {
        if (Ask-Confirmation "Recommencer ?") {
            Enable-MTRMailbox
        }
    }
}
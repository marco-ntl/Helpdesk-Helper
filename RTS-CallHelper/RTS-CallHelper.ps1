#@TODO Ajouter au path et créer shortcut pour tout le monde
#@TODO Ajouter Shortcut pour focus la window
#@TODO Ajouter onglet Workplace
using module RTS-WorkplaceAPI
using module ActiveDirectory
Add-type -AssemblyName System.Windows.Forms
#Function Start-CallHelper([bool]$makeNewProcess = $true) {
#if($makeNewProcess -eq $true){
#Start-Process "cmd" -ArgumentList '/c start', '/min ""', 'powershell.exe -noprofile -WindowStyle Hidden -Command Start-CallHelper $false'
#return
#}
#On génère déjà le token workplace pour qu'il soit placé en cache    
#Start-ThreadJob -ScriptBlock { Connect-ToWorkplace } | Out-Null
$TEMP_PASSWORDS = ("Welcome12345", "Migration2024@", "NewRTS@2024!")
$PC_LOADING = "Chargement des PCs..."
$ICON_LOCATION = "C:\Program Files\CallHelper\CallHelper.ico"
$tabIndex = 0
$script:IS_SEARCHING = $false
$pingNOKColor = [System.Drawing.Color]::Crimson
$pingOKColor = [System.Drawing.Color]::ForestGreen

$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Anchor = "Top, Left, Right"
$mainForm.Size = [System.Drawing.Point]::new(440, 468)
$mainForm.Text = "Call Helper"
$mainForm.TopMost = $true
#$mainForm.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 8.25)    

$tmrCheckJobs = [System.Windows.Forms.Timer]::new();
$tmrCheckJobs.Enabled = $false
$tmrCheckJobs.Interval = 100
$COMPUTERS_LOADING_JOB_NAME = "loadComputers"

$tclMain = New-Object System.Windows.Forms.TabControl
$tclMain.Dock = "Fill"
$tclMain.SizeMode = "Fixed"

$tabMain = New-Object System.Windows.Forms.TabPage;
$tabMain.Size = [System.Drawing.Point]::new(296, 255)
$tabMain.Text = "Users"

#Groupbox "User"
$gbxUsername = New-Object System.Windows.Forms.GroupBox
$gbxUsername.Anchor = "Top, Left, Right"
$gbxUsername.Location = [System.Drawing.Point]::new(4, 4)
$gbxUsername.Size = [System.Drawing.Size]::new(289, 54)
$gbxUsername.Text = "Recherche (Username, tél ou nom)"

$tbxUsername = New-Object System.Windows.Forms.TextBox
$tbxUsername.Anchor = "Top, Left, Right"
$tbxUsername.Location = [System.Drawing.Point]::new(9, 21)
$tbxUsername.Size = [System.Drawing.Size]::new(167, 20)
$tbxUsername.TabIndex = $tabIndex++

$btnSearchUser = New-Object System.Windows.Forms.Button
$btnSearchUser.Anchor = "Right"
#$btnSearchUser.Dock = "Right"
$btnSearchUser.Location = [System.Drawing.Point]::new(182, 19)
$btnSearchUser.Size = [System.Drawing.Size]::new(100, 24)
$btnSearchUser.Text = "Chercher"
$btnSearchUser.TabIndex = $tabIndex++

$gbxUsername.Controls.AddRange(($tbxUsername, $btnSearchUser))
$tabMain.Controls.Add($gbxUsername)

#Groupbox "Actions"
$gbxActions = New-Object System.Windows.Forms.GroupBox
$gbxActions.Anchor = "Top, Bottom, Left, Right"
$gbxActions.Location = [System.Drawing.Point]::new(4, 60)
$gbxActions.Size = [System.Drawing.Size]::new(289, 192)
$gbxActions.Text = "Actions"

$tlpActions = New-Object System.Windows.Forms.TableLayoutPanel
$tlpActions.Anchor = "Top, Bottom, Left, Right"
$tlpActions.ColumnCount = 4
$tlpActions.Location = [System.Drawing.Point]::new(10, 18)
$tlpActions.Size = [System.Drawing.Size]::new(269, 165)
#TODO ajouter marge pour scrollbar
$tlpActions.SuspendLayout()
$tlpActions.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 31)) | Out-Null
$tlpActions.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 18)) | Out-Null
$tlpActions.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 18)) | Out-Null
$tlpActions.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 33)) | Out-Null

$tlpActions.HorizontalScroll.Maximum = 0
$tlpActions.AutoScroll = $false
$tlpActions.VerticalScroll.Visible = $false;
$tlpActions.VerticalScroll.Enabled = $false
#$tlpActions.AutoScroll = $true;
#$tlpActions.VerticalScroll = $false

#Ligne 1
$lblStaticUsername = New-Object System.Windows.Forms.Label
$lblStaticUsername.Dock = "Top"
$lblStaticUsername.TextAlign = "BottomLeft"
$lblStaticUsername.Text = "Personne"
#$lblStaticUsername.AutoSize = $true
#$tlpActions.SetColumnSpan($lblStaticUsername, 1)

$pnlPerson = New-Object System.Windows.Forms.Panel
$pnlPerson.Margin = 0
$pnlPerson.Anchor = "None"

#$pnlPerson.Dock = "Fill"

$tlpActions.SetColumnSpan($pnlPerson, 2)

$lblFullName = New-Object System.Windows.Forms.Label
$lblFullName.Dock = "Top"
$lblFullName.TextAlign = "BottomCenter"
$lblFullName.Text = "Nom"

$cbxNames = New-Object System.Windows.Forms.ComboBox
$cbxNames.Dock = "Top"
$cbxNames.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cbxNames.TabIndex = $tabIndex++
$cbxNames.Visible = $false

$pnlPerson.Controls.AddRange(($lblFullName, $cbxNames))
#$lblFullName.AutoSize = $true
$tlpActions.SetColumnSpan($pnlPerson, 2)

$pnlSam = New-Object System.Windows.Forms.Panel
$pnlSam.Margin = 0
#$pnlSam.Padding.Top = 0
#$pnlSam.Padding.Bottom = 0

$lblSamAccountName = New-Object System.Windows.Forms.Label
$lblSamAccountName.Dock = "Top"
$lblSamAccountName.TextAlign = "BottomCenter"
$lblSamAccountName.Text = "Username"
#$lblSamAccountName.AutoSize = $true

$btnChoosePerson = New-Object System.Windows.Forms.Button
$btnChoosePerson.Dock = "Top"
$btnChoosePerson.Text = "Séléctionner"
$btnChoosePerson.Visible = $false
$btnChoosePerson.TabIndex = $tabIndex++

$pnlSam.Controls.AddRange(($lblSamAccountName, $btnChoosePerson))

$tlpActions.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 25)) | Out-Null
$tlpActions.Controls.AddRange(($lblStaticUsername, $pnlPerson, $pnlSam))

#Ligne 2
$lblStaticAtt14 = New-Object System.Windows.Forms.Label
$lblStaticAtt14.Dock = "Top"
$lblStaticAtt14.TextAlign = "BottomLeft"
$lblStaticAtt14.Text = "Attribut 14"
#Si autosize est set sur faux, le label ne s'alignera pas verticalement. Merci Winforms :)
#$lblStaticAtt14.AutoSize = $true

$lblAtt14 = New-Object System.Windows.Forms.Label
$lblAtt14.Dock = "Top"
$lblAtt14.TextAlign = "BottomCenter"
$lblAtt14.Text = "N/A"
#Voir commentaire au dessus
#$lblAtt14.AutoSize = $true
$tlpActions.SetColumnSpan($lblAtt14, 2)


$btnChangeAtt = New-Object System.Windows.Forms.Button
$btnChangeAtt.Dock = "Top"
$btnChangeAtt.Text = "Changer"
$btnChangeAtt.Enabled = $false
#$btnChangeAtt.Font = [System.Drawing.Font]::new("SRG SSR Type", 8.25)
#$tlpActions.SetColumnSpan($btnChangeAtt, 1)
$btnChangeAtt.TabIndex = $tabIndex++

$tlpActions.Controls.AddRange(($lblStaticAtt14, $lblAtt14, $btnChangeAtt))

#Ligne 3
$lblStaticIsLocked = New-Object System.Windows.Forms.Label
$lblStaticIsLocked.Dock = "Top"
$lblStaticIsLocked.TextAlign = "BottomLeft"
$lblStaticIsLocked.Text = "Bloqué ?"
#$lblStaticIsLocked.AutoSize = $true
#$tlpActions.SetColumnSpan($lblIsLocked, 2)

$lblIsLocked = New-Object System.Windows.Forms.Label
$lblIsLocked.Dock = "Top"
$lblIsLocked.TextAlign = "BottomCenter"
$lblIsLocked.Text = "N/A"
#$lblIsLocked.AutoSize = $true
$tlpActions.SetColumnSpan($lblIsLocked, 2)


$btnUnlock = New-Object System.Windows.Forms.Button
$btnUnlock. Dock = "Top"
$btnUnlock. Text = "Débloquer"
$btnUnlock.Enabled = $false
#$btnUnlock.Font = [System.Drawing.Font]::new("SRG SSR Type", 8.25)
$tlpActions.SetColumnSpan($btnUnlock, 1)
$btnUnlock.TabIndex = $tabIndex++

$tlpActions.Controls.AddRange(($lblStaticIsLocked, $lblIsLocked, $btnUnlock))

#Ligne 4
$lblStaticPwdAge = New-Object System.Windows.Forms.Label
$lblStaticPwdAge.Dock = "Top"
$lblStaticPwdAge.TextAlign = "BottomLeft"
$lblStaticPwdAge.Text = "Age MDP"
#$lblStaticPwdAge.AutoSize = $true
#$tlpActions.SetColumnSpan($lblStaticPwdAge, 2)

$lblPwdAge = New-Object System.Windows.Forms.Label
$lblPwdAge.Dock = "Top"
$lblPwdAge.TextAlign = "MiddleCenter"
$lblPwdAge.Text = "N/A jours"
#$lblPwdAge.AutoSize = $true
$tlpActions.SetColumnSpan($lblPwdAge, 2)

$btnGenerate = New-Object System.Windows.Forms.Button
$btnGenerate.Dock = "Top"
$btnGenerate.Text = "Générer MDP"
$btnGenerate.TabIndex = $tabIndex++

$tlpActions.Controls.AddRange(($lblStaticPwdAge, $lblPwdAge, $btnGenerate))

#Ligne 5
$lblStaticTempPwd = New-Object System.Windows.Forms.Label
$lblStaticTempPwd.Dock = "Top"
$lblStaticTempPwd.TextAlign = "BottomLeft"
$lblStaticTempPwd.Text = "MDP temp."
#$lblStaticTempPwd.AutoSize = $true

$cbxTempPwd = New-Object System.Windows.Forms.ComboBox
$cbxTempPwd. Dock = "Top"
$cbxTempPwd.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cbxTempPwd.Items.AddRange($TEMP_PASSWORDS)
$cbxTempPwd.SelectedIndex = 0
$cbxTempPwd.TabIndex = $tabIndex++
$btnShowHide.TabIndex = $tabIndex++
$tlpActions.SetColumnSpan($cbxTempPwd, 2)

$btnSetTempPwd = New-Object System.Windows.Forms.Button
$btnSetTempPwd.Dock = "Top"
$btnSetTempPwd.Text = "Changer"
$btnSetTempPwd.Enabled = $false
$btnSetTempPwd.TabIndex = $tabIndex++

$tlpActions.Controls.AddRange(($lblStaticTempPwd, $cbxTempPwd, $btnSetTempPwd))

#Ligne 6
$lblStaticNewPwd = New-Object System.Windows.Forms.Label
$lblStaticNewPwd.Dock = "Top"
$lblStaticNewPwd.TextAlign = "BottomLeft"
$lblStaticNewPwd.Text = "Nouveau MDP"
#$lblStaticNewPwd.AutoSize = $true

$tbxNewMDP = New-Object System.Windows.Forms.TextBox
$tbxNewMDP.Dock = "Top"
$tbxNewMDP.PasswordChar = '*'
$tlpActions.SetColumnSpan($tbxNewMDP, 2)
$tbxNewMDP.TabIndex = $tabIndex++

$btnShowHide = New-Object System.Windows.Forms.Button
$btnShowHide.Dock = "Top"
$btnShowHide.Text = "Afficher"

$tlpActions.Controls.AddRange(($lblStaticNewPwd, $tbxNewMDP, $btnShowHide))

#Ligne 7
$lblStaticConfirmPwd = New-Object System.Windows.Forms.Label
$lblStaticConfirmPwd.Dock = "Top"
$lblStaticConfirmPwd.TextAlign = "BottomLeft"
$lblStaticConfirmPwd.Text = "Confirmer MDP"
#$lblStaticConfirmPwd.AutoSize = $true

$tbxConfirmMDP = New-Object System.Windows.Forms.TextBox
$tbxConfirmMDP. Dock = "Top"
$tbxConfirmMDP.PasswordChar = '*'
$tlpActions.SetColumnSpan($tbxConfirmMDP, 2)
$tbxConfirmMDP.TabIndex = $tabIndex++
#$btnShowHide.TabIndex = $tabIndex++

$btnChangePwd = New-Object System.Windows.Forms.Button
$btnChangePwd.Dock = "Top"
$btnChangePwd.Text = "Changer"
$btnChangePwd.Enabled = $false
$btnChangePwd.TabIndex = $tabIndex++

$tlpActions.Controls.AddRange(($lblStaticConfirmPwd, $tbxConfirmMDP, $btnChangePwd))

#Ligne 8
$lblStaticAssignedPC = New-Object System.Windows.Forms.Label
$lblStaticAssignedPC.Dock = "Top"
$lblStaticAssignedPC.TextAlign = "BottomLeft"
$lblStaticAssignedPC.Text = "PCs assignés"
#$lblStaticAssignedPC.AutoSize = $true

$cbxAssignedPC = New-Object System.Windows.Forms.ComboBox
$cbxAssignedPC.Dock = "Top"
$cbxAssignedPC.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$tlpActions.SetColumnSpan($cbxAssignedPC, 3)
$cbxAssignedPC.TabIndex = $tabIndex++

$tlpActions.Controls.AddRange(($lblStaticAssignedPC, $cbxAssignedPC))

#Ligne 9
$lblStaticNumPC = New-Object System.Windows.Forms.Label
$lblStaticNumPC.Dock = "Top"
$lblStaticNumPC.TextAlign = "BottomLeft"
$lblStaticNumPC.Text = "N° PC"
#$lblStaticNumPC.AutoSize = $true

$tbxNumPC = New-Object System.Windows.Forms.TextBox
$tbxNumPC.Dock = "Top"
$tlpActions.SetColumnSpan($tbxNumPC, 2)
$tbxNumPC.TabIndex = $tabIndex++

$btnPing = New-Object System.Windows.Forms.Button
$btnPing.Dock = "Top"
$btnPing.Text = "Ping"
$btnPing.TabIndex = $tabIndex++

$tlpActions.Controls.AddRange(($lblStaticNumPC, $tbxNumPC, $btnPing))

#Ligne 10
$lblStaticIsOnline = New-Object System.Windows.Forms.Label
$lblStaticIsOnline.Dock = "Top"
$lblStaticIsOnline.TextAlign = "BottomLeft"
$lblStaticIsOnline.Text = "Ping ?"

$lblIsOnline = New-Object System.Windows.Forms.Label
$lblIsOnline.Dock = "Top"
$lblIsOnline.Text = "NOK"
$lblIsOnline.TextAlign = "BottomCenter"
$lblIsOnline.ForeColor = $pingNOKColor
$lblIsOnline.Font = [System.Drawing.Font]::new($lblIsOnline.Font, [System.Drawing.FontStyle]::Bold)
$tlpActions.SetColumnSpan($lblIsOnline, 2)
$lblIsOnline.TabIndex = $tabIndex

$btnRemote = New-Object System.Windows.Forms.Button
$btnRemote.Dock = "Top"
$btnRemote.Text = "Remote"
$btnRemote.TabIndex = $tabIndex++

$tlpActions.Controls.AddRange(($lblStaticIsOnline, $lblIsOnline, $btnRemote))

#Ligne 11
$btnAlwaysOnTop = New-Object System.Windows.Forms.Button
$btnAlwaysOnTop.Dock = "Top"
$btnAlwaysOnTop.Text = "Cacher la fenêtre"
$btnAlwaysOnTop.TabIndex = $tabIndex++
$tlpActions.SetColumnSpan($btnAlwaysOnTop, 2)

$lblSpacer = New-Object System.Windows.Forms.Label
$tlpActions.Controls.AddRange(($lblSpacer,$btnAlwaysOnTop))

Function check-computerPings() {
    $computer = $tbxNumPC.Text
    [System.Net.Sockets.TcpClient]::new().ConnectAsync("google.com", 80).Wait(100) 
    if ($computer.Length -gt 0 -and (Test-Connection -ComputerName $computer -Count 1 -Quiet)) {
        $lblIsOnline.ForeColor = $pingOKColor
        $lblIsOnline.Text = "OK"
    }
    else {
        $lblIsOnline.ForeColor = $pingNOKColor
        $lblIsOnline.Text = "NOK"
    }
}

Function Handle-FormLoad([object]$sender, [System.EventArgs]$e) {
    $tbxUsername.Focus()
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
        if ($groups['Local'].Success) {
            $countryCode = '+41'
        }
                    
        if ($null -eq $ndc -or $null -eq $trio -or $number -eq $firstPair -or $null -eq $secondPair -or $number -eq $countryCode) {
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

$USER_PROPERTIES = ("ExtensionAttribute14", "LockedOut", "PasswordLastSet")
Function Search-UserFuzzy([string]$data, [bool]$noPhone = $false) {
    <#
        Cette fonction essaye de trouver l'utilisateur recherché, via les étapes suivantes
        1. Check si la query correspond à un nom d'utilisateur (SAM)
        2. Check si la query correspond à un numéro de téléphone (Pour openscape)
        3. Check si la query correspond à un nom de famille
    #>
    $SEARCH_BASE = 'OU=RTS,OU=Units,DC=media,DC=int'
    <#
        Matche les numéros qui commencent par 0 (numéro Suisse) 00 (numéro étranger) + (numéro étranger)
        Il faut cependant d'abord filtrer l'input et enlever tous les charactères qui ne sont pas des chiffres (ou +)
        Format de sortie -> 0 (ou +XX) 78 123 45 67
        Groupes de capture :
            International 	-> Si le numéro est au format international, matche l'indicateur de pays(0041, +33, etc...) VIDE SI FORMAT NATIONAL (ex. 078)
            Local 			-> Contient '0' si le numéro est au format local (ex. 078) VIDE SI FORMAT INTERNATIONAL (ex. 0041, +33)
            NDC 			-> "National Destination Code", le code de l'opérateur pour les mobiles et la région pour les fixes (ex. 78, 22, 58)
            Trio 			-> Le groupe de 3 chiffres qui suit le NDC
            FirstPair 		-> La première paire, qui suit le Trio
            SecondPair		-> La deuxième paire (les deux derniers chiffres du numéro) 
        #>
    $PHONE_REGEX = '^(?:(?<International>(?:\+|00)\d{2})|(?<Local>0))(?<NDC>\d{2})(?<Trio>\d{3})(?<FirstPair>\d{2})(?<SecondPair>\d{2})'
    #On enlève les potentiels whitespaces au début et à la fin
    $data = $data.Trim()

    # ----- CHECK DATA = SAM OU LAST NAME -----
    if ($data -match "\D+") {
        $bySAM = Get-ADUser -Filter "SamAccountName -eq '$($data)'" -Properties $USER_PROPERTIES
        if ($null -ne $bySAM) {
            return $bySAM
        }

        $bySAMWildCard = Get-ADUser -Filter "SamAccountName -eq '$($data)*'" -Properties $USER_PROPERTIES
        if($null -ne $bySAMWildCard){
            return $bySAMWildCard
        }

        # ----- CHECK DATA = LAST NAME -----
        $byLastNameNoWildCard = Get-ADUser -filter "Surname -like '$($data)'" -SearchBase $SEARCH_BASE -Properties $USER_PROPERTIES
        if ($byLastNameNoWildCard -is [System.Collections.ICollection]) {
            #Si on a plus d'un résultat, inutile d'aller plus loin, la recherche par keyword prendra plus longtemps et ne trouvera rien de mieux
            return $byLastNameNoWildCard
        }
        if ($null -ne $byLastNameNoWildCard -and $byLastNameNoWildCard -isnot [System.Collections.ICollection]) {
            return $byLastNameNoWildCard
        }
        $byLastNameWildCard = Get-ADUser -filter "Surname -like '*$($data)*'" -SearchBase $SEARCH_BASE -Properties $USER_PROPERTIES

        if ($byLastNameWildCard -is [System.Collections.ICollection]) {
            #Si on a plus d'un résultat, inutile d'aller plus loin, la recherche par keyword prendra plus longtemps et ne trouvera rien de mieux
            return $byLastNameWildCard
        }

        if ($null -ne $byLastNameWildCard -and $byLastNameWildCard -isnot [System.Collections.ICollection]) {
            return $byLastNameWildCard
        }
    }

    # ----- CHECK DATA = Phone number -----
    #Ce regex matche la majorité des séparateurs utilisés dans le monde ('-', '.', ' '), les séparateurs seront ensuite strippés
    #Il permet aussi d'avoir optionnellement un + au début
    #Ne matche pas si le numéro contient des charactères invalides (lettres, charactères spéciaux hors séparateurs, etc...)
    if ($noPhone -eq $false -and $data -match "^\+?[\d\-\. ]+$") {
        $numbers = Get-PermutationsForPhoneNumber $data
        $query = (($numbers | % { "mobile -eq '$_' -or extensionAttribute14 -eq '$_' -or " }) -join '')
        $query = $query.Substring(0, $query.Length - 5) #On génère la query et on enlève le dernière ' -or ' (5 caractères)
        $user = Get-ADUser -SearchBase $SEARCH_BASE -Filter $query -Properties $USER_PROPERTIES
        if ($null -ne $user -and $user -isnot [System.Collections.ICollection]) {
            return $user
        }
    }

    # ---- DATA TYPE NOT FOUND -----
    return $false
}

Function Handle-SearchUser([object]$sender, [System.EventArgs]$e, [bool]$NoPhone = $false) {
    $global:IS_SEARCHING = $true
    #$mainForm.UseWaitCursor = $true
    $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $cbxAssignedPC.items.Clear()
    $cbxAssignedPC.Items.Add($PC_LOADING)
    $cbxAssignedPC.SelectedIndex = 0

    $user = Search-UserFuzzy $tbxUsername.Text -NoPhone:$NoPhone
    $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default

    if ($null -eq $user -or $user -eq $false) {
        $lblSamAccountName.Text = "Pas trouvé"
        $lblFullName.Text = "Pas trouvé"
        $btnUnlock.Enabled = $false
        $btnSetTempPwd.Enabled = $false
        $btnChangeAtt.Enabled = $false
        return $false
    }
    elseif ($user -is [System.Collections.ICollection]) {
        $lblSamAccountName.Text = [string]::Empty

        $lblFullName.Visible = $false;
        $lblSamAccountName.Visible = $false
        $cbxNames.Visible = $true
        $btnChoosePerson.Visible = $true

        $cbxNames.Items.Clear()
        $cbxNames.Items.AddRange(($user | % { $_.SamAccountName}))
        $cbxNames.SelectedIndex = 0
        $cbxNames.Focus()
        $cbxNames.DroppedDown = $true

        $btnUnlock.Enabled = $false
        $btnSetTempPwd.Enabled = $false
        $btnChangeAtt.Enabled = $false
        return $false
    }

    $script:user = $user
    $btnChangeAtt.Enabled = $true

    $lblFullName.Visible = $true;
    $lblSamAccountName.Visible = $true
    $cbxNames.Visible = $false
    $btnChoosePerson.Visible = $false

    $lblIsLocked.Text = if ($user.LockedOut) { "Bloqué" } else { "Pas bloqué" }
    if($null -eq $user.PasswordLastSet){ #Si PasswordLastSet est null, c'est que l'utilisateur doit changer son mot de passe à la prochaine connexion
        $lblPwdAge.Text = "0 jours"
    }else{
        $pwdAge = New-TimeSpan -Start $user.PasswordLastSet -End (Get-Date)
        $lblPwdAge.Text = "$($pwdAge.Days) jours"
    }

    $btnUnlock.Enabled = $true
    $btnSetTempPwd.Enabled = $true
    $btnChangeAtt.Enabled = $true

    if($user.ExtensionAttribute14.Length -eq 0){
        $lblAtt14.Text = "Vide"
    }else{
        $lblAtt14.Text = $user.ExtensionAttribute14
    }
    $lblSamAccountName.Text = $user.SamAccountName
    $lblFullName.Text = "$($user.Surname.ToUpper()) $($user.GivenName)"
    if ($null -ne (Get-Job $COMPUTERS_LOADING_JOB_NAME)) {
        Stop-Job $COMPUTERS_LOADING_JOB_NAME
        Remove-Job $COMPUTERS_LOADING_JOB_NAME
    }
    Start-ThreadJob { param($user) Import-Module RTS-WorkplaceAPI -Force -PassThru | Out-Null ; Get-ComputersFromUser $user } -ArgumentList $user.SamAccountName -Name $COMPUTERS_LOADING_JOB_NAME
    $tmrCheckJobs.Add_Tick( { 
            #Check-ComputersLoaded retourne $null si le job n'existe pas, et $false si il n'est pas terminé
            <#if($null -eq (#>Check-ComputersLoaded <#)){
            $tmrCheckJobs.Enabled = $false
        }#>
        })
    $tmrCheckJobs.Enabled = $true

    #if ($user.LockedOut) { $btnUnlock.Focus() } else { $tbxNewMDP.Focus() }
    $btnUnlock.Focus()
    $global:IS_SEARCHING = $false
}

$tickCounter = 0
Function Check-ComputersLoaded() {
    $tmrCheckJobs.Enabled = $false

    $job = Get-Job

    if ($job -is [System.Collections.ICollection]) {
        $job = $job | ? { $_.Name -eq $COMPUTERS_LOADING_JOB_NAME }
    }

    if ($null -eq $job) {
        #$cbxAssignedPC.Items.Clear()
        return $null
    }

    if ($job.State -eq "Completed") {
        #Le premier objet dans la liste retournée par Receive-Job contient 
        $tmrCheckJobs.Enabled = $false
        $tmrCheckJobs.Remove_Tick( { Check-ComputersLoaded })
        #$mainForm.UseWaitCursor = $false
        $data = Receive-Job $COMPUTERS_LOADING_JOB_NAME
        Stop-Job $COMPUTERS_LOADING_JOB_NAME
        Remove-Job $COMPUTERS_LOADING_JOB_NAME
        Handle-ComputersDoneLoading $data
    }
    else {
        $tmrCheckJobs.Enabled = $true
        return $false
    }
}

Function Handle-ComputersDoneLoading($computers) {
    $cbxAssignedPC.items.Clear()
    #$cbxAssignedPC.Text = ""
    $tbxNumPC.Text = ""

    if ($null -eq $computers) {
        return
    }

    if ($computers.GetType() -eq [String]) {
        $cbxAssignedPC.Items.Add($computers)
    }
    else {
        $cbxAssignedPC.Items.AddRange($computers)
    }

    if ($cbxAssignedPC.Items.Count -gt 0) {
        $cbxAssignedPC.SelectedIndex = 0
    }
}

Function Handle-UsernameTextChanged ([object]$sender, [System.EventArgs]$e) { 
    <# Désactivé car peut parfois porter à confusion
    if ($tbxUsername.TextLength -ge 6) {
        Handle-SearchUser $sender $e -NoPhone:$true
    }#>  
}

Function Handle-ChoosePersonKeyDown(){
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $btnChoosePerson.PerformClick()
    }
}

Function Handle-UsernameKeyDown([object]$sender, $e) {
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $btnSearchUser.PerformClick()
    }
}

Function Handle-UnlockClick([object]$sender, [System.EventArgs]$e) {
    if ($null -ne $script:user) {
        Unlock-ADAccount $script:user.DistinguishedName
        #On recharge le profil de l'utilisateur afin de confirmer que le compte est bien débloqué
        $btnSearchUser.PerformClick()
    }
}

Function Handle-GeneratePasswordClick([object]$sender, [System.EventArgs]$e) {
    $tbxNewMDP.PasswordChar = [Char]::MinValue
    $tbxConfirmMDP.PasswordChar = [Char]::MinValue
    $pwd = Generate-RandomPassword
    $tbxNewMDP.Text = $pwd
    $tbxConfirmMDP.Text = $pwd
    $btnShowHide.Text = "Cacher"
}

Function Handle-ShowHideClick([object]$sender, [System.EventArgs]$e) {
    if ($tbxNewMDP.PasswordChar -eq '*') {
        $tbxNewMDP.PasswordChar = [Char]::MinValue
        $tbxConfirmMDP.PasswordChar = [Char]::MinValue
        $btnShowHide.Text = "Cacher"
    }
    else {
        $tbxNewMDP.PasswordChar = '*'
        $tbxConfirmMDP.PasswordChar = '*'
        $btnShowHide.Text = "Afficher"
    }
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

Function Assert-PasswordCorrect {
    $pwdConforms = Check-PasswordConformity $tbxNewMDP.Text 
    return $pwdConforms -and ($tbxNewMDP.Text -eq $tbxConfirmMDP.Text)
}

Function Handle-PwdTextChanged([object]$sender, [System.EventArgs]$e) {
    $btnChangePwd.Enabled = Assert-PasswordCorrect
}

Function Handle-PwdTextKeyDown([object]$sender, $eventArgs) {
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter -and (Assert-PasswordCorrect)) {
        $btnChangePwd.PerformClick()
    }
}

Function Refresh-UserInfo([object]$sender, $eventArgs) {
    if ($null -eq $script:user) {
        return $null
    }
    $tbxUsername.Text = $script:user.SamAccountName
    $btnSearchUser.PerformClick()
}

Function Handle-ChangePwdClick([object]$sender, [System.EventArgs]$e) {
    if ((Assert-PasswordCorrect) -and $script:user -ne $null) {
        try {
            Set-ADAccountPassword $script:user -Reset:$true -NewPassword (ConvertTo-SecureString $tbxNewMDP.Text -AsPlainText -Force)
            #[System.Windows.Forms.MessageBox]::Show("Mot de passe changé avec succès")
            Refresh-UserInfo
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Erreur lors du changement de mot de passe, veuillez réessayer")
        }
    }
}

Function Handle-AssignedPCsChanged([object]$sender, [System.EventArgs]$e) {
    $tbxNumPC.Text = $cbxAssignedPC.SelectedItem
    Check-ComputerPings
}

Function Handle-SetTempPwdClick([object]$sender, [System.EventArgs]$e) {
    if ($null -eq $script:user) {
        return $null
    }
    Set-ADAccountPassword $script:user -Reset:$true -NewPassword (ConvertTo-SecureString $cbxTempPwd.SelectedItem -AsPlainText -Force)
    Set-ADUser $script:user -ChangePasswordAtLogon:$true
    Refresh-UserInfo
}

Function Handle-RemoteButtonClick([object]$sender, [System.EventArgs]$e) {
    Start-Process "C:\Install\5.0.8825\Files\SCCMRemoteControl5\CmRcViewer.exe" -ArgumentList $tbxNumPC.Text
}

Function Handle-RemoteTbxKeyDown([object]$sender, $eventArgs) {
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $btnPing.PerformClick()
    }
}

Function Handle-Att14BtnClicked() {
    $textInputForm = New-Object System.Windows.Forms.Form
    $textInputForm.Anchor = "Top, Left, Right"
    $textInputForm.Size = [System.Drawing.Point]::new(225, 140)
    $textInputForm.Text = "Att. 14"
    $textInputForm.MaximizeBox = $false;
    #$mainForm.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 8.25)

    $tlpActions = New-Object System.Windows.Forms.TableLayoutPanel
    $tlpActions.Anchor = "Top, Bottom, Left, Right"
    $tlpActions.ColumnCount = 2
    $tlpActions.Location = [System.Drawing.Point]::new(10, 10)
    $tlpActions.Size = [System.Drawing.Size]::new(180, 100)
    #TODO ajouter marge pour scrollbar
    #$tlpActions.SuspendLayout()
    $tlpActions.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 40)) | Out-Null
    $tlpActions.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 60)) | Out-Null
    #$tlpActions.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 18)) | Out-Null
    #$tlpActions.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 33)) | Out-Null

    $tlpActions.HorizontalScroll.Maximum = 0
    $tlpActions.AutoScroll = $false
    $tlpActions.VerticalScroll.Visible = $false;
    $tlpActions.AutoScroll = $true;

    #Ligne 1
    $lblStaticNum = New-Object System.Windows.Forms.Label
    $lblStaticNum.Dock = "Top"
    $lblStaticNum.TextAlign = "BottomLeft"
    $lblStaticNum.Text = "Num"
    #$lblStaticNum.AutoSize = $true

    $tbxNewNum = New-Object System.Windows.Forms.TextBox
    $tbxNewNum.Dock = "Top"
    $tbxNewNum.TabIndex = 0

    $tlpActions.Controls.AddRange(($lblStaticNum, $tbxNewNum))

    #Ligne 2
    $lblStaticFormatted = New-Object System.Windows.Forms.Label
    $lblStaticFormatted.Dock = "Top"
    $lblStaticFormatted.TextAlign = "BottomLeft"
    $lblStaticFormatted.Text = "Format"
    $lblStaticFormatted.AutoSize = $true

    $lblNumFormatted = New-Object System.Windows.Forms.Label
    $lblNumFormatted.Dock = "Top"
    $lblNumFormatted.TextAlign = "BottomCenter"
    $lblNumFormatted.AutoSize = $true

    $tlpActions.Controls.AddRange(($lblStaticFormatted, $lblNumFormatted))

    #Ligne 3
    $btnSetAtt = New-Object System.Windows.Forms.Button
    #$btnSetAtt.Anchor = ([System.Windows.Forms.AnchorStyles]::Top, [System.Windows.Forms.AnchorStyles]::Left, [System.Windows.Forms.AnchorStyles]::Right, [System.Windows.Forms.AnchorStyles]::Bottom)
    $btnSetAtt.Anchor = [System.Windows.Forms.AnchorStyles]::None
    $btnSetAtt.Text = "Att. 14"
    $btnSetAtt.TabIndex = 1
    
    $btnSetBoth = New-Object System.Windows.Forms.Button
    #$btnSetBoth.Anchor = ([System.Windows.Forms.AnchorStyles]::Top, [System.Windows.Forms.AnchorStyles]::Left, [System.Windows.Forms.AnchorStyles]::Right, [System.Windows.Forms.AnchorStyles]::Bottom)
    $btnSetBoth.Anchor = [System.Windows.Forms.AnchorStyles]::None
    $btnSetBoth.Text = "Att. 14 + tél"
    $btnSetBoth.TabIndex = 2

    $tlpActions.Controls.AddRange(($btnSetAtt, $btnSetBoth))

    Function Handle-TbxNumTextChanged() {
        $permutations = Get-PermutationsForPhoneNumber $tbxNewNum.Text
        if ($null -ne $permutations -and $false -ne $permutations) {
            $lblNumFormatted.ForeColor = [System.Drawing.Color]::ForestGreen
            $lblNumFormatted.Text = $permutations[0]
        }
        else {
            $lblNumFormatted.ForeColor = [System.Drawing.Color]::Crimson
            $lblNumFormatted.Text = $tbxNewNum.Text
        }
    }

    Function Get-NumValue(){
        if ($lblNumFormatted.ForeColor -eq [System.Drawing.Color]::ForestGreen) {
            return $lblNumFormatted.Text
        }
        else {
            return $tbxNewNum.Text
        }
    }

    Function Handle-ChangeAtt14Clicked() {
        $textInputForm.UseWaitCursor = $true
        $val = Get-NumValue
        Set-ADUser $script:user -Replace @{ExtensionAttribute14 = $val }
        $textInputForm.UseWaitCursor = $false
        $textInputForm.Close()
        Refresh-UserInfo
    }

   Function Handle-ChangeBothClicked() {
        $textInputForm.UseWaitCursor = $true
        $val = Get-NumValue
        Set-ADUser $script:user -Replace @{ExtensionAttribute14 = $val;Mobile = $val}
        $textInputForm.UseWaitCursor = $false
        $textInputForm.Close()
        Refresh-UserInfo
    }

    $tbxNewNum.Add_TextChanged( { Handle-TbxNumTextChanged })
    $btnSetAtt.Add_Click( { Handle-ChangeAtt14Clicked })
    $btnSetBoth.Add_Click( { Handle-ChangeBothClicked })

    $textInputForm.Controls.Add($tlpActions)
    $textInputForm.ShowDialog()

}

Function Handle-ChoosePersonClick() {
    $tbxUsername.Text = $cbxNames.SelectedItem
    $btnSearchUser.PerformClick()
}

Function Handle-AlwaysOnTop(){
    $mainForm.TopMost = !$mainForm.TopMost
    if($mainForm.TopMost){
        $btnAlwaysOnTop.Text = "Cacher la fenêtre"
    }else{
        $btnAlwaysOnTop.Text = "Afficher la fenêtre"
    }
}

$mainForm.Add_Shown( { Handle-FormLoad })

$tbxUsername.Add_TextChanged( { Handle-UsernameTextChanged } )
$tbxUsername.Add_KeyDown( { Handle-UsernameKeyDown } )
$btnSearchUser.Add_Click( { Handle-SearchUser } )

$btnChoosePerson.Add_Click( { Handle-ChoosePersonClick} )
$cbxNames.Add_KeyDown( {Handle-ChoosePersonKeyDown} )

$btnChangeAtt.Add_Click( {Handle-Att14BtnClicked} )

$btnUnlock.Add_Click( { Handle-UnlockClick })

$btnGenerate.Add_Click( { Handle-GeneratePasswordClick })
$btnShowHide.Add_Click( { Handle-ShowHideClick })

$tbxNewMDP.Add_TextChanged( { Handle-PwdTextChanged })
$tbxConfirmMDP.Add_TextChanged( { Handle-PwdTextChanged })
$tbxNewMDP.Add_KeyDown( { Handle-PwdTextKeyDown })
$tbxConfirmMDP.Add_KeyDown( { Handle-PwdTextKeyDown })

$btnChangePwd.Add_Click( { Handle-ChangePwdClick })
$btnSetTempPwd.Add_Click( { Handle-SetTempPwdClick })

$cbxAssignedPC.Add_SelectedIndexChanged( { Handle-AssignedPCsChanged })
$btnRemote.Add_Click( { Handle-RemoteButtonClick })
$tbxNumPC.Add_Keydown( { Handle-RemoteTbxKeyDown } )

$btnPing.Add_Click( { Check-ComputerPings })

$btnAlwaysOnTop.Add_Click( { Handle-AlwaysOnTop } )

$gbxActions.Controls.Add($tlpActions)
$tabMain.Controls.Add($gbxActions)

$tclMain.Controls.Add($tabMain)
$mainForm.Controls.Add($tclMain)

$tlpActions.ResumeLayout()
$mainForm.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($ICON_LOCATION)
[system.windows.forms.application]::run($mainForm)

#}

Function Compile-CallHelper() {
    ps2exe -inputFile "C:\Git\PS-Script\RTS-CallHelper\RTS-CallHelper.ps1" -outputfile "C:\Git\PS-Script\RTS-CallHelper\CallHelper.exe" -STA -noConsole -iconFile $ICON_LOCATION -title "CallHelper" -company "Radio Television Suisse" -copyright "NATALE Marco 2024" -noOutput -noError -DPIAware
}

Function Update-CallHelperExe() {
    Start-Process "cmd" -Verb runAs -ArgumentList '/c copy "C:\Git\PS-Script\RTS-CallHelper\CallHelper.exe" "C:\Program Files\CallHelper\CallHelper.exe" /Y'
}

Function Release-CallHelper() {
    Compile-CallHelper
    Update-CallHelperExe
}
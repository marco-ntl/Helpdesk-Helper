$global:TokenExpiry = (Get-Date)
$global:Session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
$global:Session.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
$defaultHeaders = @{
    "authority"          = "api.workplace.srgssr.ch"
    "method"             = "POST"
    "path"               = "/auth/login"
    "scheme"             = "https"
    "accept"             = "application/json, text/plain, */*"
    "accept-encoding"    = "gzip, deflate, br, zstd"
    "accept-language"    = "fr-FR,fr;q=0.9,en-US;q=0.8,en;q=0.7"
    "authorization"      = "Bearer null"
    "origin"             = "https://workplace.srgssr.ch"
    "referer"            = "https://workplace.srgssr.ch/"
    "sec-ch-ua"          = "`"Chromium`";v=`"122`", `"Not(A:Brand`";v=`"24`", `"Google Chrome`";v=`"122`""
    "sec-ch-ua-mobile"   = "?0"
    "sec-ch-ua-platform" = "`"Windows`""
    "sec-fetch-dest"     = "empty"
    "sec-fetch-mode"     = "cors"
    "sec-fetch-site"     = "same-site"
}
$defaultHeaders.Keys | % { $global:Session.Headers.Add($_, $defaultHeaders[$_]) }

#Il faut encapsuler les variables globales car sinon elle ne sont pas accessibles dans les Jobs
Function Get-WorkplaceURLs() {
    if ($null -eq $global:WorkplaceURLs) {
        $global:WorkplaceURLs = [PSCustomObject]@{
            BASE      =  "https://api.workplace.srgssr.ch"
            AUTH      =  "https://api.workplace.srgssr.ch/auth/login"
            COMPUTERS =  "https://api.workplace.srgssr.ch/computers"
        }
    }
    return $global:WorkplaceURLs
}

Function Get-WorkplaceArgs() {
    if ($null -eq $global:WorkplaceArgs) {
        $global:WorkplaceArgs = [PSCustomObject]@{
            USERNAME = "username"
            SEARCH = "search"
            TYPE = "type"
            MDM_TYPE = "MDM"
        }
    }
    return $global:WorkplaceArgs
}

Function Get-Session() {
    if ($global:TokenExpiry -le (Get-Date)) {
        if (Connect-ToWorkplace -eq $false) {
            return $false
        }
    }
    return $global:Session
}

Function Connect-ToWorkplace($reloadToken = $false) {
    try {
        $requestResult = Invoke-WebRequest -UseBasicParsing -Uri (Get-WorkplaceURLs).AUTH `
            -Method "POST" `
            -WebSession $global:Session `
            -ContentType "application/json" `
            -Body "{`"username`":`"username`",`"password`":`"password`"}"

        $parsedResult = $requestResult.Content | ConvertFrom-Json
        if ($null -eq $parsedResult.Token) {
            [System.Windows.Forms.MessageBox]::Show("Pas réussi à récupérer le Token de connexion à Workplace`n$($_)")
        }
        $global:TokenExpiry = (Get-Date).AddSeconds($parsedResult.expiresIn)
        $global:WorkplaceToken = $parsedResult.Token
        $global:Session.Headers.authorization = "Bearer " + $global:WorkplaceToken
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Erreur lors de la connexion à l'API Workplace`n$($_)`n$((Get-WorkplaceURLs).Auth)")
        return $false
    }
}

Function Get-ComputersFromUser([string]$username) {
    $requestResult = Invoke-WebRequest -UseBasicParsing -Uri "$((Get-WorkplaceURLs).COMPUTERS)?$((Get-WorkplaceArgs).USERNAME)=$($username)" `
        -Method "GET" `
        -WebSession (Get-Session) `
        -ContentType "application/json"
    $computers = $requestResult.Content | ConvertFrom-Json
    return $computers.computers | % { $_.DisplayName }
}

Function Get-ComputerFromSerial([string]$serial) {
    $requestResult = Invoke-WebRequest -UseBasicParsing -Uri "$((Get-WorkplaceURLs).COMPUTERS)?$((Get-WorkplaceArgs).Search)=$($serial)" `
        -Method "GET" `
        -WebSession (Get-Session) `
        -ContentType "application/json"
    $computers = $requestResult.Content | ConvertFrom-Json
    return $computers.computers | % { $_.DisplayName }
}

Function Get-MDMFromMacbook([string]$number) {
    $requestResult = Invoke-WebRequest -UseBasicParsing -Uri "$((Get-WorkplaceURLs).COMPUTERS)/$($number)?$((Get-WorkplaceArgs).Type)=$((Get-WorkplaceArgs).MDM)" `
        -Method "GET" `
        -WebSession (Get-Session) `
        -ContentType "application/json"
    $computers = $requestResult.Content.ToString().Replace("ID", "_ID") | ConvertFrom-Json
    return $computers.SCSM 
}
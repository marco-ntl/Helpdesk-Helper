using module RTS-Components

$SEARCH_MODULE = 'Recherche'
$MODULES = (("RTS-AD", "Fonctions relatives aux comptes AD"),
            ("RTS-ConferenceRooms","Fonctions relatives aux salles de conf�rence"),
            ("RTS-Dd","Fontions cr��es par David Dabo"),
            ("RTS-Outlook","Fonctions relatives aux boites mail"),
            ("RTS-Bossy","Fonctions cr��es par St�phane Bossy"),
            ("RTS-Infomut","Fonctions relatives aux infos mutations"),
            ($SEARCH_MODULE, "Chercher une fonction par nom")) | Sort-Object #On s'assure que la liste soit toujours tri�e par ordre alphab�tique


Function RTS-Menu {
    begin{
        $selectedModule = Select-Module
        $currModuleName = $MODULES[$selectedModule][0]
    }
    process{
        $commands = [System.Collections.ArrayList]::new()
    
        if($currModuleName -eq $SEARCH_MODULE){
            Write-Prompt "Recherche: "
            $query = Read-Host
            $commands = @($MODULES | % { Get-FunctionsFromModule $_[0] } | ? { $_.ToLower() -like "*$($query.ToLower())*" })
            if($commands.Length -le 0){
                Write-Err "Aucun r�sultats"
                return
            }
        }else{
            if(-not(Is-ModuleImported $currModuleName)){ #"Recherche" n'�tant pas un vrai module, il faut checker si le module est import� APR�S s'�tre assur� que ce n'est pas le module s�l�ctionn�
                Import-Module $currModuleName -WarningAction SilentlyContinue | Out-Null 
            }
            $commands = @(Get-FunctionsFromModule $currModuleName)
        }
    
        $select = [SimpleSelect]::new($currModuleName, $commands)
        $command = $commands[$select.AskUser()]
        Write-Prompt ("`n" + $command.ToUpper()) -noNewLine:$false
        Invoke-Expression $command
    }
    end{
        if(Ask-Confirmation "R�ouvrir le menu ?"){
            RTS-Menu
        }
    }
}


Function Get-FunctionsFromModule([string]$moduleName){
    return @(Get-Command -Module $moduleName | Sort-Object 'Name' |  % { return $_.Name }) #Si Get-Command ne retourne qu'une seule fonction, l'objet sera retourn� comme string et non comme un array.
                                                                                               #On ajoute donc @() pour forcer $commands � �tre un array
}

Function Select-Module {
    $choices = $MODULES | % { $_[0] + ' - ' + $_[1] }
    $select = [SimpleSelect]::new("Choix du module", $choices)
    return $select.AskUser()
}

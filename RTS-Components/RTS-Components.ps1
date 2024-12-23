class SimpleTable {
    [bool]$silent #Peut paraître inutile, mais il est plus facile de checker si l'exécution devrait être silencieuse depuis l'objet, plutôt qu'avant chaque appel de "Next" et "SetState"

    [ValidateNotNullOrEmpty()][string]$header
    [string]$headerLeftBorder
    [string]$headerRightBorder

    [ValidateNotNull()][System.Collections.ArrayList]$items
    [double]$biggestItemInTabs
    [int]$lineLength
    [int]$leftOffset
    [int]$cursor = 0
    [int]$MIN_OFFSETS = 1
    [int]$TAB_WIDTH = 8

    [string]$horizontalBox = [string][char]9552
    [string]$verticalBox = [string][char]9553
    [string]$topLeftBox = [string][char]9556
    [string]$topRightBox = [string][char]9559
    [string]$bottomLeftBox = [string][char]9562
    [string]$bottomRightBox = [string][char]9565

    [System.ConsoleColor]$borderColor = [System.ConsoleColor]::DarkRed
    [System.ConsoleColor]$headerColor = [System.ConsoleColor]::Yellow
    [System.ConsoleColor]$itemColor = [System.ConsoleColor]::Yellow


    SimpleTable([string]$header, [object[]]$items) {
        if (Test-Path variable:global:psISE) {
            #Si le script tourne dans l'ISE
            $this.TAB_WIDTH = 4 #Powershell ISE utilise des tabs de 4 charactères de long
        }
        else {
            $this.TAB_WIDTH = 8 #CMD utilise des tabs de 8 charactères de long
        }
        $this.leftOffset = 0
        $this.items = $items
        $this.biggestItemInTabs = $this.ComputeSizeInTabsWithOffset(($items | Measure-Object -Maximum -Property Length).Maximum, $true) #On récupère la taille de l'item le plus large, en tabulations
        $this.ComputeHeader($header)
    }

    [double]ComputeSizeInTabsWithOffset([int]$lengthOfLongest, [bool]$endsWithTab) {
        $nbTabs = $this.ComputeSizeInTabs('A' * $lengthOfLongest, $endsWithTab) #ComputeSizeInTabs prend un string
        return $nbTabs + $this.MIN_OFFSETS #On ajoute le nombre minimum d'offsets désirés
    }

    [int]ComputeLineLength() {
        return $this.biggestItemInTabs * $this.TAB_WIDTH + ($this.TAB_WIDTH * 2)
    }#                                                     ^    ^^^^^^^^^^^^^^^^^^^ 
    #                                                    NOK   On a une tab à gauche, et une tab à droite (la bordure gauche/droite + l'espacement = TAB_WIDTH)
    [void]SetSilent([bool]$silent) {
        $this.silent = $silent
    }

    [void]ComputeHeader([string]$header) {
        $this.header = " $header " #On ajoute un espace avant et après le header, plus esthétique
        $sizeInTabs = $this.ComputeSizeInTabsWithOffset($this.header.Length, $false)
        if ($this.biggestItemInTabs -lt $sizeInTabs) {
            $this.biggestItemInTabs = $sizeInTabs
        }
        $this.lineLength = $this.ComputeLineLength()
        #On aligne les bordures du header avec la checklist
        $halfBorderLength = ($this.lineLength - $this.header.Length - 2) / 2.0 #Longueur du header -2 car le coin gauche et droit n'utiliseront pas le même charactère que la ligne
        $this.headerLeftBorder = $this.topLeftBox + ($this.horizontalBox * [System.Math]::Ceiling($halfBorderLength)) #On alterne Floor et ceiling, au cas ou la longueur de halfborder n'est pas paire
        $this.headerRightBorder = ($this.horizontalBox * [System.Math]::Floor($halfBorderLength)) + $this.topRightBox #Cela permet de d'éviter de désaligner l'en-tête
    }

    [void]ShowHeader() {
        if ($this.silent -eq $false) {
            Write-Host #On Affiche une newline avant
            Write-Host $this.headerLeftBorder -NoNewline -f $this.borderColor
            Write-Host $this.header -NoNewline -f $this.headerColor
            Write-Host $this.headerRightBorder -f $this.borderColor
        }
    }

    [void]Show() {
        $this.ShowHeader()
        $this.ShowCurrent()
        (2..$this.items.Count) | % {$this.Next()} #2..items.length == items.Length - 2; On retire 1 car le tableau est indexé à 0, et 1 car le premier item est déjà affiché
        $this.cursor = 0 #On reset le tableau, comme ça on peut le réafficher sans avoir à l'initialiser à nouveau
    }

    [double]ComputeSizeInTabs([string]$text, [bool]$endsWithTab) {
        if ($endsWithTab -eq $true) {
            return [System.Math]::Ceiling($text.Length / $this.TAB_WIDTH) #Si $text sera terminé par une tab, on fait Ceiling car la tab qui termine $text occupera la place manquante
        }
        else {
            return [System.Math]::Floor($text.Length / $this.TAB_WIDTH) #Sinon, on fait Floor
        }
    }
    
    [void]FatalError() {
        $this.SetState($false)
        if($this.cursor -lt $this.items.Count - 1){
            $this.ShowBottom()
        }
    }

    [void]ShowCurrent() {
        if ($this.silent -eq $false) {
            Write-Host "$($this.verticalBox)`t" -NoNewline -f $this.borderColor #Si on ne mets pas une tabulation, les éléments du tableau ne commenceront pas alignés sur une colonne
            Write-Host (("`t" * $this.leftOffset) + $this.items[$this.cursor]) -NoNewLine -f $this.itemColor #ce qui peut poser des soucis au niveau des calculs
            if ($this.items[$this.cursor].Length % $this.TAB_WIDTH -ne 0) {
                #Si l'item ne finit pas aligné sur une colonne, on rajoute une tab à la fin pour aligner le curseur
                Write-Host "`t" -NoNewline
            }
            $nbOffsets = ($this.biggestItemInTabs - $this.ComputeSizeInTabs($this.items[$this.cursor], $true))
            Write-Host $("`t" * $nbOffsets) -NoNewline
            Write-Host (" " * ($this.TAB_WIDTH - 1)) -NoNewline #L'espace vide à gauche du tableau : |`t, càd TAB_WIDTH - 1 (car on a un caractère, puis une tabulation)
            Write-Host "$($this.verticalBox)" -f $this.borderColor #On met donc un espacement de TAB_WIDTH - 1 à droite aussi, afin que l'espacement soit symmétrique des deux côtés

            if ($this.cursor -ge $this.items.Count - 1) {
                #Si on vient d'afficher le dernier item, on affiche aussi le bas du tableau
                $this.ShowBottom()
            }
        }
    }

    [void]ShowBottom() {
        Write-Host ($this.bottomLeftBox + ($this.horizontalBox * ($this.lineLength - 2)) + $this.bottomRightBox) -f $this.borderColor
        Write-Host #On ajoute une ligne d'espacement en dessous
    }

    [void]Next() {
        if($this.cursor -lt $this.items.Count - 1){
            $this.cursor++
            $this.ShowCurrent()
        }
    }

    [void]Inject($item, $offset = 0) {
        if($this.cursor -eq 0){
            $this.items = $($item, $this.items)
            return
        }

        $this.items.Insert($this.cursor + $offset + 1, $item) # = $($this.items[0..$this.cursor + $offset]; $item ; $this.items[($this.cursor + $offset + 1)..($this.items.Length - 1)]) #Insert un item dans la prochaine position de $items
    }

}

class CheckList : SimpleTable {

    CheckList([string]$name,[object[]]$choices) : base($name,$choices){
        
    }

    [void]Start(){
        $this.ShowHeader()
        $this.ShowCurrent()
    }

    [int]ComputeLineLength() {
        return $this.biggestItemInTabs * $this.TAB_WIDTH + 3 + ($this.TAB_WIDTH * 2)
    }#                                                     ^    ^^^^^^^^^^^^^^^^^^^ 
    #                                                    NOK   On a une tab à gauche, et une tab à droite (la bordure gauche/droite + l'espacement = TAB_WIDTH)
    [void]ShowCurrent() {
        if ($this.silent -eq $false) {
            Write-Host "$($this.verticalBox)`t" -NoNewline -f $this.borderColor #Si on ne mets pas une tabulation, les éléments du tableau ne commenceront pas alignés sur une colonne
            Write-Host (("`t" * $this.leftOffset) + $this.items[$this.cursor]) -NoNewLine -f $this.itemColor #ce qui peut poser des soucis au niveau des calculs
            if ($this.items[$this.cursor].Length % $this.TAB_WIDTH -ne 0) {
                #Si item ne finit pas aligné sur une colonne, on rajoute une tab à la fin pour aligner le curseur
                Write-Host "`t" -NoNewline
            }
        }
    }

    [void]SetState([bool]$state) {
        if ($this.silent -eq $false) {
            $nbOffsets = ($this.biggestItemInTabs - $this.ComputeSizeInTabs($this.items[$this.cursor], $true))
            Write-Host $("`t" * $nbOffsets) -NoNewline
            if ($state) {
                Write-Host " " -NoNewline #On ajoute un espace avant, afin que "OK" fasse la même taille que "NOK", ce qui facilite l'alignement; On doit l'ajouter à la ligne d'avant, sinon l'espace récupèrera le style de Write-Success
                Write-Success "OK" -noNewLine:$true
            }
            else {
                Write-Err "NOK" -noNewLine:$true -extraNewLine:$false
            }
            Write-Host (" " * ($this.TAB_WIDTH - 1)) -NoNewline #L'espace vide à gauche du tableau : |`t, càd TAB_WIDTH - 1 (car on a un caractère, puis une tabulation)
            Write-Host "$($this.verticalBox)" -f $this.borderColor #On met donc un espacement de TAB_WIDTH - 1 à droite aussi, afin que l'espacement soit symmétrique des deux côtés
            if ($this.cursor -ge ($this.items.Count - 1)) {
                #Si on vient d'afficher le dernier item, on affiche aussi le bas du tableau
                $this.ShowBottom()
            }
        }
    }

    [void]SetStateAndGoToNext([bool]$state) {
        $this.SetState($state)
        $this.Next()
    }

}

class SimpleSelect : SimpleTable {

    SimpleSelect([string]$name, [string[]]$items) : base($name, $this.AddIndexes($items)) {
    }
    
    [string[]] AddIndexes($choices){
        return $choices | % {$i = 1} { "$i. $_";$i++}        
    }

    [int]AskUser(){
        $this.Show()
        $choice = 0
        while($choice -le 0 -or $choice -gt $this.items.Count){
            Write-Prompt "Séléction : "
            try{
            $choice = [int]::Parse((Read-Host))
            }catch{
                continue; #Si l'entrée de l'utilisateur n'est pas un nombre, on revient au début de la boucle
            }
        }
        return ($choice - 1) #On retourne l'index de l'élément choisi
    }
}

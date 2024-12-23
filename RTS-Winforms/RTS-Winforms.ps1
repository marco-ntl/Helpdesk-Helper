Add-Type -AssemblyName System.Windows.Forms

$MIN_SEARCH_LENGTH = 3
enum CONTROLS {
    Search = 0
    Textbox = 1
    Select = 2
    Date = 3
    Listbox = 4
    Combobox = 5
}

class Acronyms {
    static $Label = 'lbl'
    static $Textbox = 'tbx'
    static $Listbox = 'lbx'
    static $Combobox = 'cbx'
    static $RadioButton = 'rdb'
    static $DatePicker = 'dtp'
    static $TabControl = 'tbc'
}

class ControlArgs {
    [string]$label
    [string]$name
}

class TextControlArgs : ControlArgs {
    [string]$format = '.*' #@TODO créer fonction "Check-Format(regex,text)
    [System.Drawing.Color]$backgroundColor
    [System.Drawing.Color]$foregroundColor
    [bool]$multiline
}

class TextboxArgs : TextControlArgs { }

class SearchArgs : TextControlArgs {
    [System.Func[string, System.Array]]$predicate
}

class DateArgs : ControlArgs {
    [DateTime]$start = [DateTime]::MinValue
    [DateTime]$end = [DateTime]::MaxValue
}

class ListControlArgs : ControlArgs{
    [System.Array]$choices
    [bool]$allowEmpty
    [bool]$multipleChoices
}

class ListboxArgs : ListControlArgs { }

class ComboboxArgs : ListControlArgs { }

class RadioSelectArgs : ListControlArgs { }

class OutputConsoleArgs : TextControlArgs { }

Function Create-Form([string]$title,[System.Drawing.Size]$size){
    $form = New-Object System.Windows.Forms.Form

    $form.ClientSize = [System.Drawing.Size]::new(500,300)
    $form.text = $title
    return $form
}

Function Make-Search($arguments) {
    $label = New-Object System.Windows.Forms.Label
    $textbox = New-Object System.Windows.Forms.TextBox
    $listbox = New-Object System.Windows.Forms.ListBox

    $label.name = [Acronyms]::Label + $arguments.name
    $textbox.name = [Acronyms]::Textbox + $arguments.name
    $listbox.name = [Acronyms]::Listbox + $arguments.name

    $label.Text = $arguments.label

    $textbox.Add_Keyup({param($sender, $e)  #Passer arguments par arguments
        if($sender.text.Length -lt $MIN_SEARCH_LENGTH){
            return
        }

        $predicateResult = Invoke-Command $arguments.predicate -ArgumentList $sender.text
        $listbox.Items.Clear()
        $listbox.Items.AddRange($predicateResult)
    })
    $listbox.Location.Y += $textbox.Size.Height
    $listbox.Size.Width = $textbox.Size.Width
    return ($label,($textbox,$listbox))
}

Function Make-Listbox([ListboxArgs]$args) {
    $label = New-Object System.Windows.Forms.Label
    $listbox = New-Object System.Windows.Forms.ListBox

    $listbox.Name = [Acronyms]::Listbox + $args.name
    $label.name = [Acronyms]::Label + $args.label

    $label.Text = $args.label
    $listbox.Items.AddRange($args.choices)

    return ($label,$listbox)
}

Function Make-Textbox([TextboxArgs]$args) {
    $label = New-Object System.Windows.Forms.Label
    $textbox = New-Object System.Windows.Forms.TextBox

    $label.Text = [Acronyms]::Label + $args.label
    $textbox.Name = [Acronyms]::Textbox + $args.Name

    $label.text = $args.label

    return ($label, $textbox)
}



Function test{ #Prototype Control "Search"
    [System.Windows.Forms.Form]$form = Create-Form 'test'
    $cbx = New-Object System.Windows.Forms.ComboBox
    $cbx.Name = 'test'
    $cbx.Add_KeyUp({param($sender,$e)yoyoyo $sender $e})
    $form.Controls.Add($cbx)
    $form.ShowDialog()
}

function yoyoyo([object]$sender, [System.Windows.Forms.KeyEventArgs]$e){ #Prototype control "Search
    if($sender.Text.Length -le 3)
    {
        return
    }
    $u = get-aduser -Filter "SamAccountName -like '$($sender.Text)*'"
    if($null -ne $u){
        $u | % { $sender.Items.Add($_.SamAccountName)}
    }
}
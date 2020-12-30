Import-Module "$PSScriptRoot\Modules\New-FilePermissionGroups.psm1"
Import-Module "$PSScriptRoot\Modules\Convert-SerbianToEnglish.psm1"
Import-Module "$PSScriptRoot\Modules\Write-Log.psm1"

$Credential = Get-Credential

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$MainForm                        = New-Object system.Windows.Forms.Form
$MainForm.ClientSize             = New-Object System.Drawing.Point(800,359)
$MainForm.text                   = "Create New Permission Groups"
$MainForm.TopMost                = $false
$MainForm.Icon                   = ".\group.ico"

$FolderPathLabel                 = New-Object system.Windows.Forms.Label
$FolderPathLabel.text            = "Folder Path"
$FolderPathLabel.AutoSize        = $true
$FolderPathLabel.width           = 25
$FolderPathLabel.height          = 10
$FolderPathLabel.location        = New-Object System.Drawing.Point(25,35)
$FolderPathLabel.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$OUPathLabel                     = New-Object system.Windows.Forms.Label
$OUPathLabel.text                = "OU Path"
$OUPathLabel.AutoSize            = $true
$OUPathLabel.width               = 25
$OUPathLabel.height              = 10
$OUPathLabel.location            = New-Object System.Drawing.Point(25,80)
$OUPathLabel.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$FolderPathTextBox               = New-Object system.Windows.Forms.TextBox
$FolderPathTextBox.multiline     = $false
$FolderPathTextBox.width         = 650
$FolderPathTextBox.height        = 20
$FolderPathTextBox.location      = New-Object System.Drawing.Point(120,30)
$FolderPathTextBox.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$OUPathTextBox                   = New-Object system.Windows.Forms.TextBox
$OUPathTextBox.multiline         = $false
$OUPathTextBox.width             = 650
$OUPathTextBox.height            = 20
$OUPathTextBox.location          = New-Object System.Drawing.Point(120,75)
$OUPathTextBox.Font              = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$CreateGroupsButton              = New-Object system.Windows.Forms.Button
$CreateGroupsButton.text         = "Create Groups"
$CreateGroupsButton.width        = 140
$CreateGroupsButton.height       = 30
$CreateGroupsButton.location     = New-Object System.Drawing.Point(330,305)
$CreateGroupsButton.Font         = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$ResultTextBox                   = New-Object system.Windows.Forms.TextBox
$ResultTextBox.multiline         = $true
$ResultTextBox.text              = "Waiting for operation execution"
$ResultTextBox.width             = 746
$ResultTextBox.height            = 158
$ResultTextBox.enabled           = $false
$ResultTextBox.location          = New-Object System.Drawing.Point(25,120)
$ResultTextBox.Font              = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$MainForm.controls.AddRange(@($FolderPathLabel,$OUPathLabel,$FolderPathTextBox,$OUPathTextBox,$CreateGroupsButton,$ResultTextBox))

$CreateGroupsButton.Add_Click({
    New-FilePermissionGroups -OUPath $OUPathTextBox.text -FolderPath $FolderPathTextBox.text -Credential $Credential |
    ForEach-Object {
        $ResultTextBox.Text = $_
    }
})

[void]$MainForm.ShowDialog()
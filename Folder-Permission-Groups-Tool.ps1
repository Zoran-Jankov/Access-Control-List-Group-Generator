<#
.SYNOPSIS
Creates file permissions Read Only and Read-Write AD groups for a shared folder, and grants them appropriate access to the shared
folder.

.DESCRIPTION
Creates file permissions Read Only and Read-Write AD groups for a shared folder, and grants them appropriate access to the share
folder. It names AD groups by appending folder name to the prefix `PG-RO-` for AD group that has Read Only access and `PG-RW-` for
the AD group that has Read-Write access. It generates log for events and error.

.NOTES
Version:        1.2
Author:         Zoran Jankov
#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

$FolderPermissionGroupsOU = "OU=File Server Permission Groups"

$Credential = Get-Credential
$RootOU = $FolderPermissionGroupsOU + "," + (Get-ADDomain).DistinguishedName

if (-not ([adsi]::Exists("LDAP://$RootOU"))) {
    New-ADOrganizationalUnit -Name $FolderPermissionGroupsOU -Path (Get-ADDomain).DistinguishedName
}

$OrganisationalUnits = [ordered]@{}
Get-ADOrganizationalUnit -SearchBase $RootOU -SearchScope Subtree -Filter * | ForEach-Object {
    $OrganisationalUnits.Add($_.Name, $_.DistinguishedName)
    }

$LogTitle = "********************************************************  Folder Permission Groups Tool Log  *********************************************************"
$LogSeparator = "******************************************************************************************************************************************************"
    
#-----------------------------------------------------------[Functions]------------------------------------------------------------

<#
.SYNOPSIS
Writes a log entry to console, log file and report file.

.DESCRIPTION
Creates a log entry with timestamp and message passed thru a parameter Message, and saves the log entry to log file, to report log
file, and writes the same entry to console. In "Settings.cfg" file paths to report log and permanent log file are contained, and
option to turn on or off whether a console output, report log and permanent log should be written. If "Settings.cfg" file is absent
it loads the default values. Depending on the NoTimestamp parameter, log entry can be written with or without a timestamp.
Format of the timestamp is "yyyy.MM.dd. HH:mm:ss:fff", and this function adds " - " after timestamp and before the main message.

.PARAMETER Message
A string message to be written as a log entry

.PARAMETER NoTimestamp
A switch parameter if present timestamp is disabled in log entry

.EXAMPLE
Write-Log -Message "A log entry"

.EXAMPLE
Write-Log "A log entry"

.EXAMPLE
Write-Log -Message "===========" -NoTimestamp

.EXAMPLE
"A log entry" | Write-Log

.NOTES
Version:        2.2
Author:         Zoran Jankov
#>
function Write-Log {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true,
                   Position = 0,
                   ValueFromPipeline = $true,
                   ValueFromPipelineByPropertyName = $true,
                   HelpMessage = "A string message to be written as a log entry")]
        [string]
        $Message,

        [Parameter(Mandatory = $false,
                   Position = 1,
                   ValueFromPipeline = $true,
                   ValueFromPipelineByPropertyName = $true,
                   HelpMessage = "A switch parameter if present timestamp is disabled in log entry")]
        [switch]
        $NoTimestamp = $false
    )

    begin {
        if (Test-Path -Path ".\Settings.cfg") {
            $Settings = Get-Content ".\Settings.cfg" | ConvertFrom-StringData

            $LogFile         = $Settings.LogFile
            $ReportFile      = $Settings.ReportFile
            $WriteTranscript = $Settings.WriteTranscript -eq "true"
            $WriteLog        = $Settings.WriteLog -eq "true"
            $SendReport      = $Settings.SendReport -eq "true"
        }
        else {
            $Desktop = [Environment]::GetFolderPath("Desktop")
            $LogFile         = "$Desktop\Log.log"
            $ReportFile      = "$Desktop\Report.log"
            $WriteTranscript = $true
            $WriteLog        = $true
            $SendReport      = $false
        }
        if (-not (Test-Path -Path $LogFile)) {
            New-Item -Path $LogFile -ItemType File
        }
        if ((-not (Test-Path -Path $ReportFile)) -and $SendReport) {
            New-Item -Path $ReportFile -ItemType File
        }
    }

    process {
        if (-not($NoTimestamp)) {
            $Timestamp = Get-Date -Format "yyyy.MM.dd. HH:mm:ss:fff"
            $LogEntry = "$Timestamp - $Message"
        }
        else {
            $LogEntry = $Message
        }

        if ($WriteTranscript) {
            Write-Verbose $LogEntry -Verbose
        }
        if ($WriteLog) {
            Add-content -Path $LogFile -Value $LogEntry
        }
        if ($SendReport) {
            Add-content -Path $ReportFile -Value $LogEntry
        }
    }
}

<#
.SYNOPSIS

.DESCRIPTION
Long description

.PARAMETER InitialDirectory
Initial directory to be opend with folder browser dialog

.EXAMPLE
Get-Folder -InitialDirectory "D:\"

.EXAMPLE
Get-Folder "D:\"

.NOTES
Version:        1.0
Author:         Zoran Jankov
#>
function Get-Folder {
    [CmdletBinding()]
    param (
    [Parameter(Mandatory = $true,
                   Position = 0,
                   ValueFromPipeline = $true,
                   ValueFromPipelineByPropertyName = $true,
                   HelpMessage = "Initial directory to be opend with folder browser dialog")]
        [string]
        $InitialDirectory
    )
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    $FolderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $FolderBrowserDialog.RootFolder = 'MyComputer'

    if ($InitialDirectory) {
        $FolderBrowserDialog.SelectedPath = $InitialDirectory
    }
    [void] $FolderBrowserDialog.ShowDialog()

    return $FolderBrowserDialog.SelectedPath
}

<#
.SYNOPSIS
Creates file permissions AD groups for a shared folder

.DESCRIPTION
Creates file permissions Read Only and Read-Write AD groups for a shared folder, and grants them appropriate access to the shared folder.

.PARAMETER OUPath
Organization unit path for the permission groups

.PARAMETER FolderPath
Full path of the shared folder

.PARAMETER Credential
Domanin credential for creation of an Active Directory group

.EXAMPLE
New-FilePermissionGroups -OUPath "OU=File Server Permission Groups,DC=company,DC=com" -FolderPath "\\SERVER\Shared_Folder"

.NOTES
Version:        1.3
Author:         Zoran Jankov
#>
function New-FilePermissionGroups {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [Parameter(Mandatory = $true,
                   Position = 0,
                   ValueFromPipeline = $false,
                   ValueFromPipelineByPropertyName = $true,
                   HelpMessage = "Organization unit path for the permission groups")]
        [string]
        $OUPath,

        [Parameter(Mandatory = $true,
                   Position = 1,
                   ValueFromPipeline = $false,
                   ValueFromPipelineByPropertyName = $true,
                   HelpMessage = "Full path of the shared folder")]
        [string]
        $FolderPath,

        [Parameter(Mandatory = $true,
                   Position = 2,
                   ValueFromPipeline = $true,
                   ValueFromPipelineByPropertyName = $true,
                   HelpMessage = "Credential for creation of an Active Directory group")]
        [System.Management.Automation.PSCredential]
        $Credential
    )

    process {
        $Result = ""

        if (-not (Test-Path -Path $FolderPath)) {
            $Message = "ERROR - $FolderPath folder does not exists"
            Write-Log -Message $Message
            $Result += "$Message`r`n"
            Write-Output -InputObject $Result
            break
        }
        if (-not ([adsi]::Exists("LDAP://$OUPath"))) {
            $Message = "ERROR - $OUPath organizational unit does not exists"
            Write-Log -Message $Message
            $Result += "$Message`r`n"
            Write-Output -InputObject $Result
            break
        }
        $BaseName = (Split-Path -Path $FolderPath -Leaf).Trim()
        $Groups = @(
            @{
                Access = "ReadAndExecute"
                Prefix = "PG-RO-"
            }
            @{
                Access = "Modify"
                Prefix = "PG-RW-"
            }
        )
        foreach ($Group in $Groups) {
            $Name = $Group.Prefix + $BaseName
            try {
                New-ADGroup -Name $Name `
                            -DisplayName $Name `
                            -Path $OUPath `
							-GroupCategory Security `
							-GroupScope Global `
                            -Description $FolderPath `
                            -Credential $Credential
            }
            catch {
                $Message = "Failed to create $Name AD group `r`n" + $_.Exception
                Write-Log -Message $Message
                $Result += "$Message`r`n"
                Write-Output -InputObject $Result
                break
            }
            $Message = "Successfully created $Name AD group"
            Write-Log -Message $Message
            $Result += "$Message`r`n"
            $ACL = Get-ACL -Path $FolderPath
            $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule($Name, $Group.Access, 'ContainerInherit, ObjectInherit', 'None', 'Allow')
            try {
                $ACL.SetAccessRule($AccessRule)
                $ACL | Set-Acl -Path $FolderPath
            }
            catch {
                $Message = "Failed to grant " + $Group.Access + " access to $Name ADGroup to $FolderPath `r`n" + $_.Exception
                Write-Log -Message $Message
                $Result += "$Message`r`n"
                continue
            }
            $Message = "Successfully granted " + $Group.Access + " access to $Name ADGroup to ""$FolderPath"" shared folder"
            Write-Log -Message $Message
            $Result += "$Message`r`n"
        }
        Write-Output -InputObject $Result
    }
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$MainForm                        = New-Object system.Windows.Forms.Form
$MainForm.ClientSize             = New-Object System.Drawing.Point(800,359)
$MainForm.text                   = "Create New Permission Groups"
$MainForm.TopMost                = $false
$MainForm.FormBorderStyle        = 'Fixed3D'
$MainForm.MaximizeBox            = $false

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
$FolderPathTextBox.width         = 530
$FolderPathTextBox.height        = 20
$FolderPathTextBox.location      = New-Object System.Drawing.Point(120,30)
$FolderPathTextBox.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$CreateGroupsButton              = New-Object system.Windows.Forms.Button
$CreateGroupsButton.text         = "Create Groups"
$CreateGroupsButton.width        = 140
$CreateGroupsButton.height       = 30
$CreateGroupsButton.location     = New-Object System.Drawing.Point(330,305)
$CreateGroupsButton.Font         = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$ResultTextBox                   = New-Object system.Windows.Forms.TextBox
$ResultTextBox.multiline         = $true
$ResultTextBox.text              = "Waiting for operation execution"
$ResultTextBox.width             = 745
$ResultTextBox.height            = 160
$ResultTextBox.ScrollBars        = "Vertical"
$ResultTextBox.location          = New-Object System.Drawing.Point(25,120)
$ResultTextBox.Font              = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$ResultTextBox.ReadOnly          = $true
$ResultTextBox.ForeColor         = [System.Drawing.ColorTranslator]::FromHtml("#7ed321")
$ResultTextBox.BackColor         = [System.Drawing.ColorTranslator]::FromHtml("#000000")

$OUPathComboBox                  = New-Object system.Windows.Forms.ComboBox
$OUPathComboBox.width            = 650
$OUPathComboBox.height           = 5
$OUPathComboBox.location         = New-Object System.Drawing.Point(120,75)
$OUPathComboBox.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$SelectFolderButton              = New-Object system.Windows.Forms.Button
$SelectFolderButton.text         = "Select Folder"
$SelectFolderButton.width        = 100
$SelectFolderButton.height       = 30
$SelectFolderButton.location     = New-Object System.Drawing.Point(672,25)
$SelectFolderButton.Font         = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$MainForm.controls.AddRange(@(
    $FolderPathLabel,
    $OUPathLabel,
    $FolderPathTextBox,
    $CreateGroupsButton,
    $ResultTextBox,
    $OUPathComboBox,
    $SelectFolderButton
))

$SelectFolderButton.Add_Click({
    $FolderPathTextBox.text = Get-Folder -InitialDirectory "D:\"
})

$FolderPathTextBox.Add_Click({
    $FolderPathTextBox.text = Get-Folder -InitialDirectory "D:\"
})

$CreateGroupsButton.Add_Click({
    Write-Log -Message $LogTitle -NoTimestamp
    Write-Log -Message $LogSeparator -NoTimestamp
    $OUPath = $OrganisationalUnits.Get_Item($OUPathComboBox.Text)
    New-FilePermissionGroups -OUPath $OUPath -FolderPath $FolderPathTextBox.text -Credential $Credential |
    ForEach-Object {
        $ResultTextBox.Text = $_
    }
    Write-Log -Message $LogSeparator -NoTimestamp
})

foreach ($Item in $OrganisationalUnits.Keys.GetEnumerator()) {
    $OUPathComboBox.Items.Add($Item)
}

[void]$MainForm.ShowDialog()
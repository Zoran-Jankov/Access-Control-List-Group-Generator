<#
.SYNOPSIS
Creates file permissions Read Only and Read-Write AD groups for a shared folder, and grants them appropriate access to the shared
folder.

.DESCRIPTION
Creates file permissions Read Only and Read-Write AD groups for a shared folder, and grants them appropriate access to the share
folder. It names AD groups by appending folder name to the prefix `PG-RO-` for AD group that has Read Only access and `PG-RW-` for
the AD group that has Read-Write access. It generates log for events and error.

.NOTES
Version:        1.3
Author:         Zoran Jankov
#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

$OrganisationalUnits = [ordered]@{}
Get-ADOrganizationalUnit -SearchScope Subtree -Filter * | ForEach-Object {
    $OrganisationalUnits.Add($_.DistinguishedName)
}

$LogTitle = "********************************************************  Folder Permission Groups Tool Log  *********************************************************"
$LogSeparator = "******************************************************************************************************************************************************"
    
#-----------------------------------------------------------[Functions]------------------------------------------------------------

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
        $Desktop = [Environment]::GetFolderPath("Desktop")
        $LogFile = "$Desktop\Log.log"
    }

    process {
        if (-not($NoTimestamp)) {
            $Timestamp = Get-Date -Format "yyyy.MM.dd. HH:mm:ss:fff"
            $LogEntry = "$Timestamp - $Message"
        }
        else {
            $LogEntry = $Message
        }
        Add-content -Path $LogFile -Value $LogEntry
    }
}

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
        $FolderPath
    )

    process {
        $Result = ""
        if (-not (Test-Path -Path $FolderPath)) {
            $Message = "ERROR - '$FolderPath' folder does not exists"
            Write-Log -Message $Message
            $Result += "$Message`r`n"
            Write-Output -InputObject $Result
            break
        }
        if (-not ([adsi]::Exists("LDAP://$OUPath"))) {
            $Message = "ERROR - '$OUPath' organizational unit does not exists"
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
                    -GroupScope DomainLocal `
                    -Description $FolderPath
            }
            catch {
                $Message = "Failed to create '$Name' AD group `r`n" + $_.Exception
                Write-Log -Message $Message
                $Result += "$Message`r`n"
                Write-Output -InputObject $Result
                break
            }
            $Message = "Successfully created '$Name' AD group"
            Write-Log -Message $Message
            $Result += "$Message`r`n"
            $ACL = Get-ACL -Path $FolderPath
            $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule($Name, $Group.Access, 'ContainerInherit, ObjectInherit', 'None', 'Allow')
            try {
                $ACL.SetAccessRule($AccessRule)
                $ACL | Set-Acl -Path $FolderPath
            }
            catch {
                $Message = "Failed to grant '" + $Group.Access + "' access to '$Name' ADGroup to '$FolderPath' `r`n" + $_.Exception
                Write-Log -Message $Message
                $Result += "$Message`r`n"
                continue
            }
            $Message = "Successfully granted '" + $Group.Access + "' access to '$Name' ADGroup to '$FolderPath' shared folder"
            Write-Log -Message $Message
            $Result += "$Message`r`n"
        }
        Write-Output -InputObject $Result
    }
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$MainForm = New-Object system.Windows.Forms.Form
$MainForm.ClientSize = New-Object System.Drawing.Point(800, 560)
$MainForm.Text = "Folder Permission Groups Tool"
$MainForm.TopMost = $true
$MainForm.FormBorderStyle = 'Fixed3D'
$MainForm.MaximizeBox = $false
$MainForm.ShowIcon = $false
$MainForm.ControlBox = $false

$FolderPathLabel = New-Object system.Windows.Forms.Label
$FolderPathLabel.Text = "Folder Path"
$FolderPathLabel.AutoSize = $true
$FolderPathLabel.Width = 25
$FolderPathLabel.Height = 10
$FolderPathLabel.Location = New-Object System.Drawing.Point(25, 35)
$FolderPathLabel.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)

$OUPathLabel = New-Object system.Windows.Forms.Label
$OUPathLabel.Text = "OU Path"
$OUPathLabel.AutoSize = $true
$OUPathLabel.Width = 25
$OUPathLabel.Height = 10
$OUPathLabel.Location = New-Object System.Drawing.Point(25, 80)
$OUPathLabel.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)

$FolderPathTextBox = New-Object system.Windows.Forms.TextBox
$FolderPathTextBox.Multiline = $false
$FolderPathTextBox.Width = 530
$FolderPathTextBox.Height = 20
$FolderPathTextBox.Location = New-Object System.Drawing.Point(120, 30)
$FolderPathTextBox.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)

$CreateGroupsButton = New-Object system.Windows.Forms.Button
$CreateGroupsButton.Text = "Create Groups"
$CreateGroupsButton.Width = 140
$CreateGroupsButton.Height = 30
$CreateGroupsButton.Location = New-Object System.Drawing.Point(330, 505)
$CreateGroupsButton.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$ResultTextBox = New-Object system.Windows.Forms.TextBox
$ResultTextBox.Multiline = $true
$ResultTextBox.Text = "Waiting for operation execution..."
$ResultTextBox.Width = 745
$ResultTextBox.Height = 360
$ResultTextBox.ScrollBars = "Vertical"
$ResultTextBox.Location = New-Object System.Drawing.Point(25, 120)
$ResultTextBox.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
$ResultTextBox.ReadOnly = $true
$ResultTextBox.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#00FF41")
$ResultTextBox.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#0D0208")

$OUPathComboBox = New-Object system.Windows.Forms.ComboBox
$OUPathComboBox.Width = 650
$OUPathComboBox.Height = 5
$OUPathComboBox.Location = New-Object System.Drawing.Point(120, 75)
$OUPathComboBox.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
$OUPathComboBox.AutoCompleteMode = 'SuggestAppend'
$OUPathComboBox.AutoCompleteSource = 'ListItems'

$SelectFolderButton = New-Object system.Windows.Forms.Button
$SelectFolderButton.Text = "Select Folder"
$SelectFolderButton.Width = 100
$SelectFolderButton.Height = 30
$SelectFolderButton.Location = New-Object System.Drawing.Point(672, 25)
$SelectFolderButton.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)

$MainForm.controls.AddRange(@(
        $FolderPathLabel,
        $OUPathLabel,
        $FolderPathTextBox,
        $CreateGroupsButton,
        $ResultTextBox,
        $OUPathComboBox,
        $SelectFolderButton
    ))

$SelectFolderButton.Add_Click( {
        $FolderPathTextBox.Text = Get-Folder -InitialDirectory "D:\"
    })

$FolderPathTextBox.Add_Click( {
        $FolderPathTextBox.Text = Get-Folder -InitialDirectory "D:\"
    })

$CreateGroupsButton.Add_Click( {
        Write-Log -Message $LogTitle -NoTimestamp
        Write-Log -Message $LogSeparator -NoTimestamp
        $OUPath = $OrganisationalUnits.Get_Item($OUPathComboBox.Text)
        New-FilePermissionGroups -OUPath $OUPath -FolderPath $FolderPathTextBox.text - $ |
        ForEach-Object {
            $ResultTextBox.Text = $_
        }
        Write-Log -Message $LogSeparator -NoTimestamp
    })

foreach ($Item in $OrganisationalUnits.Keys.GetEnumerator()) {
    $OUPathComboBox.Items.Add($Item)
}

[void]$MainForm.ShowDialog()
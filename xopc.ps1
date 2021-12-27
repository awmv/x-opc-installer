$ErrorActionPreferenceBak = $ErrorActionPreference
$ErrorActionPreference = 'Stop'

$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if ($principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    if (!(Get-Module -ListAvailable -Name powershell-yaml)) {
        try {
            if ((GET-Culture).Name -eq "de-CH" -or "de-DE") {
                Install-Module powershell-yaml -Confirm:$True -Force
            }
            else {
                Install-Module powershell-yaml
            }
        }
        catch {
            Write-Host "Something went wrong installting module powershell-yaml. Perpans check if you're online."
        }
    }


    # Stole that GUI from somewhere :O
    function Read-MultiLineInputBoxDialog([string]$Message, [string]$WindowTitle, [string]$DefaultText) {
        Add-Type -AssemblyName System.Drawing
        Add-Type -AssemblyName System.Windows.Forms
     
        # Create the Label.
        $label = New-Object System.Windows.Forms.Label
        $label.Location = New-Object System.Drawing.Size(10, 10) 
        $label.Size = New-Object System.Drawing.Size(280, 20)
        $label.AutoSize = $true
        $label.Text = $Message
     
        # Create the TextBox used to capture the user's text.
        $textBox = New-Object System.Windows.Forms.TextBox 
        $textBox.Location = New-Object System.Drawing.Size(10, 40) 
        $textBox.Size = New-Object System.Drawing.Size(575, (200 + 150))
        $textBox.AcceptsReturn = $true
        $textBox.AcceptsTab = $false
        $textBox.Multiline = $true
        $textBox.ScrollBars = 'Both'
        $textBox.Text = $DefaultText
        $textBox.Font = "Consolas"
     
        # Create the OK button.
        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Location = New-Object System.Drawing.Size(415, (250 + 150))
        $okButton.Size = New-Object System.Drawing.Size(75, 25)
        $okButton.Text = "OK"
        $okButton.Add_Click( { $form.Tag = $textBox.Text; $form.Close() })
     
        # Create the Cancel button.
        $cancelButton = New-Object System.Windows.Forms.Button
        $cancelButton.Location = New-Object System.Drawing.Size(510, (250 + 150))
        $cancelButton.Size = New-Object System.Drawing.Size(75, 25)
        $cancelButton.Text = "Cancel"
        $cancelButton.Add_Click( { $form.Tag = $null; $form.Close() })
     
        # Create the form.
        $form = New-Object System.Windows.Forms.Form 
        $form.Text = $WindowTitle
        $form.Size = New-Object System.Drawing.Size(610, (320 + 150))
        $form.FormBorderStyle = 'FixedSingle'
        $form.StartPosition = "CenterScreen"
        $form.AutoSizeMode = 'GrowAndShrink'
        $form.Topmost = $True
        $form.AcceptButton = $okButton
        $form.CancelButton = $cancelButton
        $form.ShowInTaskbar = $true
     
        # Add all of the controls to the form.
        $form.Controls.Add($label)
        $form.Controls.Add($textBox)
        $form.Controls.Add($okButton)
        $form.Controls.Add($cancelButton)
     
        # Initialize and show the form.
        $form.Add_Shown( { $form.Activate() })
        $form.ShowDialog() > $null   # Trash the text of the button that was clicked.
     
        # Return the text that the user entered.
        return $form.Tag
    }

    $placeHolder = @'
install:
- name: OPC_DLS_1
  padt: 25136
  systemid: 10027
- name: OPC_DLS_2
  padt: 25137
  systemid: 10028
- name: OPC_DLS_3
  padt: 25137
  systemid: 10028
- name: OPC_DLS_4
  padt: 25137
  systemid: 10028
- name: OPC_MDS_1
  padt: 25137
  systemid: 10028

# Double check this. You can cancel to skip this step
# This window will reopen on Exception
'@

    While ($true) {
        try {
            $multiLineText = Read-MultiLineInputBoxDialog -Message "Please enter YAML-compliant configuration" -WindowTitle "Install X-OPC Services" -DefaultText $placeHolder
            $yaml = ConvertFrom-Yaml $multiLineText
            if (!$multiLineText) { 
                Write-Host "You clicked Cancel"
                break
            }
            else {
                if ($yaml.install) {
                    $yaml.install | Select-Object @{ Label = "Name"; Expression = { $_["name"] } }, @{ Label = "PADT"; Expression = { $_["padt"] } }, @{ Label = "SystemID"; Expression = { $_["systemid"] } } | ForEach-Object {
                        # Write-Host("serviceName: $($_.name)", "PADT: $($_.padt)", "systemId: $($_.systemid)", "action: install")
                        & "$PSScriptRoot\service.ps1" -serviceName $($_.name) -PADT $($_.padt) -systemid $($_.systemid) -action "install"
                    }
                }
                if ($yaml.uninstall) {
                    $yaml.uninstall | Select-Object @{ Label = "Name"; Expression = { $_["name"] } }, @{ Label = "PADT"; Expression = { $_["padt"] } }, @{ Label = "SystemID"; Expression = { $_["systemid"] } } | ForEach-Object {
                        # Write-Host("serviceName: $($_.name)", "PADT: $($_.padt)", "systemId: $($_.systemid)", "action: uninstall")
                        & "$PSScriptRoot\service.ps1" -serviceName $($_.name) -PADT $($_.padt) -systemid $($_.systemid) -action "uninstall"
                    }
                }
                if ($yaml.install -or $yaml.uninstall) {
                    break
                }
            }
        }
        catch {
            Write-Host "Something wrong with the configuration. Please try again.`n Exception:`r`n $_.Exception.Message"
            Start-Sleep -Seconds 1
        }
        finally {
            $ErrorActionPreference = $ErrorActionPreferenceBak
        }
    }
}
else {
    Start-Process -FilePath "powershell" -ArgumentList "$('-File ""')$(Get-Location)$('\')$($MyInvocation.MyCommand.Name)$('""')" -Verb runAs
}

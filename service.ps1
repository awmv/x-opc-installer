#Requires -Version 4.0
#Requires -RunAsAdministrator

param (
    [Parameter(Mandatory)][ValidatePattern("^OPC_(?:DLS|MDS)_\d{1}$")][string]$serviceName,
    [Parameter(Mandatory)][ValidatePattern("^\d{5}$")][int32]$PADT,
    [Parameter(Mandatory)][ValidatePattern("^\d{5}$")][int32]$systemId,
    [Parameter(Mandatory)][ValidatePattern("^install|uninstall$")][string]$action
)

$path = "HKLM:\SOFTWARE\WOW6432Node\HIMA\X-OPC\$serviceName"
$dataDir = Join-Path -Path 'C:\ProgramData\HIMA\X-OPC\' -ChildPath $serviceName
$exeDir = Join-Path -Path 'C:\Program Files (x86)\HIMA\X-OPC' -ChildPath $serviceName

function CreateKeyPath {
    Write-Host "CreateKeyPath"
    if (Test-Path $path) {
        $key = Get-Item -Path $path
    }
    else {
        $key = New-Item -Path $path -Force
    }
    return $key
}

function AddToRegistry {
    Write-Host "AddToRegistry"
    $key = CreateKeyPath
    $AE = '{' + [guid]::NewGuid().ToString() + '}'
    $DA = '{' + [guid]::NewGuid().ToString() + '}'
    New-ItemProperty -Path $key.PSPath -Name CLSID_AE -Value $AE | Out-Null
    New-ItemProperty -Path $key.PSPath -Name CLSID_DA -Value $DA | Out-Null
    New-ItemProperty -Path $key.PSPath -Name DataDir -PropertyType ExpandString -Value $dataDir | Out-Null
    New-ItemProperty -Path $key.PSPath -Name PADT_Port1 -PropertyType DWord -Value $PADT | Out-Null
    New-ItemProperty -Path $key.PSPath -Name PADT_Port2 -PropertyType DWord -Value 00000 | Out-Null
    New-ItemProperty -Path $key.PSPath -Name SizePendingRequests -PropertyType DWord -Value 10000 | Out-Null
    New-ItemProperty -Path $key.PSPath -Name SystemId -PropertyType DWord -Value $systemId | Out-Null
}

function RemoveFromRegistry {
    Write-Host "RemoveFromRegistry"
    if (Test-Path -Path $path) {
        Remove-Item -Path $path -Force -Verbose | Out-Null
    }
}

function CreateDirectoryInProgramFiles {
    Write-Host "CreateDirectoryInProgramFiles"
    if (![System.IO.Directory]::Exists($exeDir)) { [System.IO.Directory]::CreateDirectory($exeDir) | Out-Null }
}

function CreateBinaries {
    Write-Host "CreateBinaries"
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    function Unzip {
        param([string]$zipfile, [string]$outpath)
    
        [System.IO.Compression.ZipFile]::ExtractToDirectory($zipfile, $outpath)
    }
    Unzip "C:\Temp\hima_service.zip" $exeDir | Out-Null
}

function RemoveBinaries {
    Write-Host "RemoveBinaries"
	if (![System.IO.Directory]::Exists($exeDir)) {
		return
	}
    if ((GET-Culture).Name -eq "de-CH" -or "de-DE") {
        & takeown /f $exeDir /r /d J | Out-Null
        Remove-Item -Recurse -Force $exeDir | Out-Null
        return
    }
    & takeown /f $exeDir /r /d Y | Out-Null
    Remove-Item -Recurse -Force $exeDir | Out-Null
}

function RegisterService {
    Write-Host "RegisterService"
    & (Join-Path -Path $exeDir -ChildPath "X-OPC.exe") /regserver=$serviceName | Out-Null 
}

function UnregisterService {
    Write-Host "UnregisterService"
	if (!(Test-Path -Path (Join-Path -Path $exeDir -ChildPath "X-OPC.exe"))) {
		return
	}
    & (Join-Path -Path $exeDir -ChildPath "X-OPC.exe") /unregserver=$serviceName | Out-Null
}

function RemoveProgramData {
    Write-Host "RemoveProgramData"
	if ([System.IO.Directory]::Exists($exeDir)) {
		Write-Host "$exeDir does not exist"
		return
	}
    if ((GET-Culture).Name -eq "de-CH" -or "de-DE") {
        & ICACLS $dataDir /q /c /t /reset | Out-Null
		& takeown /f $dataDir /r /a /d J | Out-Null
		Remove-Item -Recurse -Force $dataDir | Out-Null
		return
	}
    & ICACLS $dataDir /q /c /t /reset | Out-Null
	& takeown /f $dataDir /r /a /d J | Out-Null
	Remove-Item -Recurse -Force $dataDir | Out-Null
}


function DeleteService {
    Write-Host "DeleteService"
	
    & "sc.exe" delete $serviceName | Out-Null
}


function StartService {
    Write-Host "StartService"
    Set-Service -Name $serviceName -StartupType Automatic | Out-Null
    Start-Service -Name $serviceName | Out-Null
}

function install {
    Write-Host(">Installing $serviceName")
    AddToRegistry
    CreateDirectoryInProgramFiles
    CreateBinaries
    RegisterService
    StartService  
}

function uninstall {
    Write-Host(">Uninstalling $serviceName")
    UnregisterService
    RemoveFromRegistry
    RemoveBinaries
    DeleteService
    RemoveProgramData
}

if ($action -eq "install") {
    install
	C:\Temp\BGInfo\Bginfo64.exe C:\Temp\Bginfo.bgi /TIMER:00 /SILENT /NOLICPROMPT
}

elseif ($action -eq "uninstall") {
    uninstall
	C:\Temp\BGInfo\Bginfo64.exe C:\Temp\Bginfo.bgi /TIMER:00 /SILENT /NOLICPROMPT
}

Write-Host ">All done"

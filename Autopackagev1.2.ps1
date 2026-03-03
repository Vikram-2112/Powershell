
#========================================================================
#Variable Declaration
#========================================================================
try{
$ErrorActionPreference = "Stop"
$Script:PrimaryPath     = (Get-ChildItem "$PSScriptRoot\PrimaryInstaller" | ? {$_.name -like "*.msi"}).FullName
$Script:SecondaryPaths = (Get-ChildItem "$PSScriptRoot\SecondaryInstallers" | ? {$_.name -like "*.msi"}).FullName
}
catch{

  write-host "failed to get primary or secondary installer paths: $Error[0] " -ForegroundColor Red

}
[string]$script:SInstallCMDs =   @()
[string]$Script:SUninstallCMDs = @()

#========================================================================
#Functions to get details of MSIs
#========================================================================

function Get-MsiVersion {

    param (

        [Parameter(Mandatory)]

        [string]$Path

    )


    $installer = New-Object -ComObject WindowsInstaller.Installer

    $database = $installer.OpenDatabase($Path, 0)

    $view = $database.OpenView("SELECT Value FROM Property WHERE Property = 'ProductVersion'")

    $view.Execute()

    $record = $view.Fetch()

    $version = $record.StringData(1)

    $view.Close()


    return $version

}

function Get-MSIProductname {

    param (

        [Parameter(Mandatory)]

        [string]$Path

    )


    $installer = New-Object -ComObject WindowsInstaller.Installer

    $database = $installer.OpenDatabase($Path, 0)

    $view = $database.OpenView("SELECT Value FROM Property WHERE Property = 'ProductName'")

    $view.Execute()

    $record = $view.Fetch()

    $Productname = $record.StringData(1)

    $view.Close()


    return $Productname

}
 
function Get-MSIPublisher {

    param (

        [Parameter(Mandatory)]

        [string]$Path

    )


    $installer = New-Object -ComObject WindowsInstaller.Installer

    $database = $installer.OpenDatabase($Path, 0)

    $view = $database.OpenView("SELECT Value FROM Property WHERE Property = 'Manufacturer'")

    $view.Execute()

    $record = $view.Fetch()

    $Publisher = $record.StringData(1)

    $view.Close()


    return $Publisher

} 

function Get-MSIProdcode {

    param (

        [Parameter(Mandatory)]

        [string]$Path

    )


    $installer = New-Object -ComObject WindowsInstaller.Installer

    $database = $installer.OpenDatabase($Path, 0)

    $view = $database.OpenView("SELECT Value FROM Property WHERE Property = 'ProductCode'")

    $view.Execute()

    $record = $view.Fetch()

    $Prodcode = $record.StringData(1)

    $view.Close()


    return $Prodcode

}

#==========================================================================
#For primary installer get all details.
#==========================================================================

if ($Script:PrimaryPath -ne $null){
[string]$productversion = (Get-MsiVersion -Path $Script:PrimaryPath)
$productversion = $productversion -replace '\s',''
[string]$Productcode = (Get-MSIProdcode -Path $Script:PrimaryPath)
$Productcode = $Productcode -replace '\s',''
[string]$productname = (Get-MSIProductname -Path $Script:PrimaryPath)
$productname = $productname -replace '\s',''
[string]$productvendor = (Get-MSIPublisher -Path $Script:PrimaryPath)
$productvendor = $productvendor -replace '\s',''
 

$PinstallName = "$productvendor"+"_$productname"+"_$productversion"+"x86x64"
$PinstallTitle = "$productvendor"+" "+"$productname"+" "+"$productversion"+" "+"x86x64"
$PLogName = "$productvendor"+"_$productname"+"_$productversion"+"x86x64"+"_Script.log" 

#Install command for Primary MSI Installer
Write-Host "Generating install command for primary installer"
$Pbasename = (Get-ChildItem "$PSScriptRoot\PrimaryInstaller" | ? {$_.name -like "*.msi"}).basename
$PInstallCMD += @' 
Execute-MSI -Action 'Install' -Path "{0}.MSI" -Parameters "/qn" -LogName "{1}"
'@ -f $PBasename, $Plogname

Write-Host "Install command for primary installer is: `n$PInstallCMD " -ForegroundColor Green

$PUnInstallCMD += @' 
Execute-MSI -Action 'UnInstall' -Path "{0}" -Parameters "/qn" -LogName "{1}"
'@ -f $Productcode, $("Unisntall_"+$Plogname)

Write-Host "Uninstall command for primary installer is: `n$PUnInstallCMD" -ForegroundColor Green
}

#=============================================================================
#Reteieve Secondary Installers details.
#=============================================================================
if($SecondaryPaths -ne $null){
Foreach($paths in $Script:SecondaryPaths ){
        
        Write-Host "Getting details for secondary installers from $paths" -ForegroundColor Green
          
        $SBaseName = (Get-ItemProperty $paths).Basename
        [String]$logname = "$(Get-MSIProductname $paths)"+"$(Get-MsiVersion $paths)"+'.log'
        write-host "Generated logname for secandary installer: $logname"

        $Script:SInstallCMDs += @' 
         Execute-MSI -Action 'Install' -Path "{0}.MSI" -Parameters "/qn" -LogName "{1}"
'@ -f $SBaseName, $logname
        
        Write-Host "Generated install command for $SBasename"

        [string]$Script:SUninstallCMDs += @'
        Execute-MSI -Action 'UnInstall' -Path "{0}" -Parameters "/qn" -LogName "{1}"

'@ -f $([string](Get-MSIProdcode -Path $paths)), $("uninstall_"+$logname)
        Write-Host "Generated Uninstall command for $SBasename"

        }

        Write-Host "Install commands for secondary installers:`n$SInstallCMDs" -ForegroundColor Magenta  
        Write-host "Uninstall Commands for Secondaries: `n$SUninstallCMDs"
}


#=================================================================================
#Create PSADT v3 installation script
#=================================================================================
if(($Script:PrimaryPath -ne $null) -or ($Script:SecondaryPaths -ne $null)){

$PSADTcode = Get-Content -Path "$PSScriptRoot\supportfiles\WithSecondary.cfg"

{$PSADTcode}.Invoke() | Out-File "$PSScriptRoot\Toolkit\Deploy-Application.ps1"
Copy-Item "$PSScriptRoot\PrimaryInstaller\*" -Destination "$PSScriptRoot\Toolkit\files"
Copy-Item "$PSScriptRoot\SecondaryInstallers\*" -Destination "$PSScriptRoot\Toolkit\files"
}

#=========================================================================================
#Create Intunewin file
#=========================================================================================

if(($Script:PrimaryPath -ne $null) -or ($Script:SecondaryPaths -ne $null)){
                
                try{
                write-host "Trying to generate .Intunewin file aka Intune Package" -ForegroundColor Yellow
                Start-Process -FilePath "$PSScriptRoot\IntuneWinAppUtility\IntuneWinAppUtil.exe" -ArgumentList "-c `"$PSScriptRoot\Toolkit`" -s `"$PSScriptRoot\Toolkit\Deploy-Application.exe`" -o `"$PSScriptRoot\IntunePackage`" -q" -NoNewWindow -Wait -ErrorAction Stop
                }
                catch{
                      Write-Host "Failed to generate Intune package: $error[0]" -ForegroundColor Red
                     }
                }
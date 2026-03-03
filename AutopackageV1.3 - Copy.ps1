#========================================================================
#Adding Script Parameters
#========================================================================
      [CmdletBinding(DefaultParameterSetName = 'None')]
      Param(
       #[Parameter(mandatory, ParameterSetName='EXE')]
      
      
       
       [Parameter(mandatory, ParameterSetName='EXE')]
       #[ValidateSet('MSI','EXE','APPX')]
       [switch]$EXE,

       [Parameter(mandatory, ParameterSetName='MSI')]
       [switch]$MSI,

       #[Parameter(Mandatory, ParameterSetName='Install')]
       [Parameter(mandatory = $true, HelpMessage="For EXE installers you have to define the application name for this parameter", ParameterSetName='EXE')]
       [String]$AppName,
       [Parameter(mandatory = $true, HelpMessage="For EXE installers you have to define the application version for this parameter", ParameterSetName='EXE')]
       [Version]$AppVersion ,
       [Parameter(mandatory = $true, HelpMessage="For EXE installers you have to define the application Manufacturer for this parameter", ParameterSetName='EXE')]
       [String]$AppVendor,       
       [String]$InstallArguments,
       [String]$Architechture,
       [string]$UnInstallArg

       
 
       )

#=============================================================================
#Creating visual message function
#=============================================================================
    Add-Type -AssemblyName System.Windows.Forms
    Function Show-message {


    Param(
    [String]$message,
    [validateset("Information", "Warning", "Error")]
    [String]$Type
    )
    
    [System.Windows.Forms.MessageBox]::Show(
    $message,
    $Type,
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::$Type
)
}

#=========================================================================================
#Create Intunewin file
#=========================================================================================


 Function Create-Intunewin{
                try{
                write-host "Trying to generate .Intunewin file aka Intune Package" -ForegroundColor Yellow
                Start-Process -FilePath "$PSScriptRoot\IntuneWinAppUtility\IntuneWinAppUtil.exe" -ArgumentList "-c `"$PSScriptRoot\Toolkit`" -s `"$PSScriptRoot\Toolkit\Deploy-Application.exe`" -o `"$PSScriptRoot\IntunePackage`" -q" -NoNewWindow -Wait -ErrorAction Stop
                }
                catch{
                      Write-Host "Failed to generate Intune package: $error[0]" -ForegroundColor Red
                     }
      }

 #========================================================================
 #Remove any exsting files
 #=======================================================================

 $ExistingFileCheck = Get-ChildItem -Path "$PSScriptRoot\Toolkit\Files" -ErrorAction Continue
 if ($ExistingFileCheck.count -ne 0){
    
    Remove-Item -Path "$PSScriptRoot\Toolkit\Files\*" -Force -Recurse

 }


 if($EXE){


 #========================================================
 #Prepare .exe installation script.
 #========================================================
 $script:ChkExe = Get-ChildItem "$PSScriptRoot\PrimaryInstaller" | ? {$_.name -like "*.exe"}
 if ($script:ChkExe.count -le 0){
 
      Show-message -message "No exe installer found at: $PSScriptRoot\PrimaryInstaller" -Type Error
      
      }

 else{

 $Script:ExeName     = (Get-ChildItem "$PSScriptRoot\PrimaryInstaller" | ? {$_.name -like "*.exe"}).basename

 $PinstallName = "$AppVendor"+"_$AppName"+"_$AppVersion"+"_$Architechture"
 $PinstallTitle ="$AppVendor"+"_$AppName"+"_$AppVersion" 
 
 if(!$InstallArguments){
 $InstallCMDs = @' 
   Start-Process -FilePath "$PSScriptRoot\files\{0}.exe" -ArgumentList "/s" -Wait
'@ -f $ExeName
  }
  else{
  $InstallCMDs = @'
  Start-Process -FilePath "$PSScriptRoot\files\{0}.exe" -ArgumentList {1} -Wait
'@ -f $ExeName, $InstallArguments
  }

  if ($UnInstallArg -ne $null){
   
   $UnInstallCMD = @'
  Start-Process -FilePath "$PSScriptRoot\files\{0}.exe" -ArgumentList {1} -Wait
'@ -f $ExeName, $UnInstallArg
   
   }
   else{
  $UnInstallCMD = @'
  Start-Process -FilePath "$PSScriptRoot\files\{0}.exe" -ArgumentList {1} -Wait
'@ -f $ExeName, "/uninstall /s "
   }
 $replacements = @{
                  
                '%%appVendor%%'             = $AppVendor
                '%%appName%%'               = $AppName
                '%%appVersion%%'            = $AppVersion
                '%%installName%%'           = $PinstallName
                '%%installTitle%%'          = $PinstallTitle
                #'%%SecondaryInstallers%%'   = $SInstallCMDs
                '%%ExeInstallerCMD%%'       = $InstallCMDs
                '%%ExeUninstallCMD%%'       = $UnInstallCMD
                #'%%SecondaryUninstallCMD%%' = $SUninstallCMDs
                #'%%RepairCMD%%'             = $PRepairCMD
                  
                  
                  }
    
    $Template = Get-Content -Path "$PSScriptRoot\supportfiles\ExeInstaller.Template" -Raw

foreach ($key in $replacements.Keys) {
    $template = $template.Replace($key, $replacements[$key])
}

$Template |  Out-File "$PSScriptRoot\Toolkit\Deploy-Application.ps1" -Encoding utf8

Copy-Item "$PSScriptRoot\PrimaryInstaller\*.exe" -Destination "$PSScriptRoot\Toolkit\files" 
#Copy-Item "$PSScriptRoot\SecondaryInstallers\*" -Destination "$PSScriptRoot\Toolkit\files"

Create-Intunewin

       }
}



elseif($MSI){

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

$PRepairCMD += @' 
Execute-MSI -Action 'Repair' -Path "{0}" -Parameters "/qn" -LogName "{1}"
'@ -f $Productcode, $("Repair_"+$Plogname)

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

$replacements = @{
                  
                '%%appVendor%%'             = $productvendor
                '%%appName%%'               = $productname
                '%%appVersion%%'            = $productversion
                '%%installName%%'           = $PinstallName
                '%%installTitle%%'          = $PinstallTitle
                '%%SecondaryInstallers%%'   = $SInstallCMDs
                '%%PrimaryInstallers%%'     = $PInstallCMD
                '%%PrimaryUninstallCMD%%'   = $PUnInstallCMD
                '%%SecondaryUninstallCMD%%' = $SUninstallCMDs
                '%%RepairCMD%%'             = $PRepairCMD
                  
                  
                  }


$Template = Get-Content -Path "$PSScriptRoot\supportfiles\WithSecondary.cfg" -Raw

foreach ($key in $replacements.Keys) {
    $template = $template.Replace($key, $replacements[$key])
}

$Template |  Out-File "$PSScriptRoot\Toolkit\Deploy-Application.ps1" -Encoding utf8

Copy-Item "$PSScriptRoot\PrimaryInstaller\*" -Destination "$PSScriptRoot\Toolkit\files"
Copy-Item "$PSScriptRoot\SecondaryInstallers\*" -Destination "$PSScriptRoot\Toolkit\files"

Create-Intunewin

}
}



else{
$msg = @"
You have selected, unsupported type of installer or haven't selected any of the installer type!
Please define installer tye using parameters
Available Parameters: -EXE, -MSI
"@
Show-message -message $msg -Type Error
}
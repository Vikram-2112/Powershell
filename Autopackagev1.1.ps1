
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

$PSADTcode = ( ( @'


[CmdletBinding()]

    Param (

    [Parameter(Mandatory = $false)]

    [ValidateSet('Install', 'Uninstall', 'Repair')]

    [String]$DeploymentType = 'Install',

    [Parameter(Mandatory = $false)]

    [ValidateSet('Interactive', 'Silent', 'NonInteractive')]

    [String]$DeployMode = 'Interactive',

    [Parameter(Mandatory = $false)]

    [switch]$AllowRebootPassThru = $false,

    [Parameter(Mandatory = $false)]

    [switch]$TerminalServerMode = $false,

    [Parameter(Mandatory = $false)]

    [switch]$DisableLogging = $false

)

 

Try {

    ## Set the script execution policy for this process

    Try {

        Set-ExecutionPolicy -ExecutionPolicy 'ByPass' -Scope 'Process' -Force -ErrorAction 'Stop'

    }

    Catch {

    }

 

    ##*===============================================

    ##* VARIABLE DECLARATION

    ##*===============================================

    ## Variables: Application

'@)+"`r`n"+(
@'

    [String]$appVendor = "{0}"

    [String]$appName = "{1}" 

    [String]$appVersion = "{2}"

    [String]$appArch = 'x86x64'

    [String]$appLang = 'EN'

    [String]$appRevision = '01'

    [String]$appScriptVersion = '1.0.0'

    [String]$appScriptDate = '07/09/2025'

    [String]$appScriptAuthor = 'Deloitte Technology Support, Deloitte Touche Tohmatsu'

    ##*===============================================

    ## Variables: Install Titles (Only set here to override defaults set by the toolkit)

    [String]$installName = "{3}"

    [String]$installTitle = "{4}"

    [String]$LogName = "{5}"

'@ -f $productvendor, $productname, $productversion, $PinstallName, $PinstallTitle, $PLogName) +"`r`n" +(

@'

    ##* Do not modify section below

    #region DoNotModify

 

    ## Variables: Exit Code

    [Int32]$mainExitCode = 0

 

    ## Variables: Script

    [String]$deployAppScriptFriendlyName = 'Deploy Application'

    [Version]$deployAppScriptVersion = [Version]'3.9.3'

    [String]$deployAppScriptDate = '02/05/2023'

    [Hashtable]$deployAppScriptParameters = $PsBoundParameters

 

    ## Variables: Environment

    If (Test-Path -LiteralPath 'variable:HostInvocation') {

        $InvocationInfo = $HostInvocation

    }

    Else {

        $InvocationInfo = $MyInvocation

    }

    [String]$scriptDirectory = Split-Path -Path $InvocationInfo.MyCommand.Definition -Parent

 

    ## Dot source the required App Deploy Toolkit Functions

    Try {

        [String]$moduleAppDeployToolkitMain = "$scriptDirectory\AppDeployToolkit\AppDeployToolkitMain.ps1"

        If (-not (Test-Path -LiteralPath $moduleAppDeployToolkitMain -PathType 'Leaf')) {

            Throw "Module does not exist at the specified location [$moduleAppDeployToolkitMain]."

        }

        If ($DisableLogging) {

            . $moduleAppDeployToolkitMain -DisableLogging

        }

        Else {

            . $moduleAppDeployToolkitMain

        }

    }

    Catch {

        If ($mainExitCode -eq 0) {

            [Int32]$mainExitCode = 60008

        }

        Write-Error -Message "Module [$moduleAppDeployToolkitMain] failed to load: `n$($_.Exception.Message)`n `n$($_.InvocationInfo.PositionMessage)" -ErrorAction 'Continue'

        ## Exit the script, returning the exit code to SCCM

        If (Test-Path -LiteralPath 'variable:HostInvocation') {

            $script:ExitCode = $mainExitCode; Exit

        }

        Else {

            Exit $mainExitCode

        }

    }

 

    #endregion

    ##* Do not modify section above

    ##*===============================================

    ##* END VARIABLE DECLARATION

    ##*===============================================

 

    If ($deploymentType -ine 'Uninstall' -and $deploymentType -ine 'Repair') {

        ##*===============================================

        ##* PRE-INSTALLATION

        ##*===============================================

        [String]$installPhase = 'Pre-Installation'

 

        <## Check if MSICache folder is hidden if not set it to hidden

        $MSICacheFolder = (Get-Item -Path "$env:SystemDrive\MSICache" -Force)

        If ($MSICacheFolder.Attributes -NotLike '*Hidden*') {$MSICacheFolder.Attributes+='Hidden'}

        If ($MSICacheFolder.Attributes -Like '*System*') {$MSICacheFolder.Attributes-='System'}

 
        #>
        ## <Perform Pre-Installation tasks here>

 

 

        ##*===============================================

        ##* INSTALLATION

        ##*===============================================

        [String]$installPhase = 'Installation'

 

        ## <Perform Installation tasks here>

       

        ## Installation Command for Windows Installer (.msi)

'@
)+"`r`n" +(

@'     

{0}
{1}

'@ -f $script:SInstallCMDs, $PInstallCMD 
)+"`r`n" +(

@'

        ##*===============================================

        ##* POST-INSTALLATION

        ##*===============================================

        [String]$installPhase = 'Post-Installation'

 

        ## <Perform Post-Installation tasks here>

                            

        ## Display a message at the end of the install

    }

    ElseIf ($deploymentType -ieq 'Uninstall') {

        ##*===============================================

        ##* PRE-UNINSTALLATION

        ##*===============================================

        [String]$installPhase = 'Pre-Uninstallation'

 

        ## <Perform Pre-Uninstallation tasks here>

 

        ##*===============================================
'@
)+"`r`n"+
(
@'
        ##* UNINSTALLATION

        ##*===============================================

        [String]$installPhase = 'Uninstallation'

 

        ## <Perform Uninstallation tasks here>

 

        ## Uninstallation Command for Windows Installer (.msi)
'@)+"`r`n" +(
@'
       {0}
       {1}
'@ -f $SUninstallCMDs, $PUnInstallCMD )+"`r`n" +(
@'
        ##* POST-UNINSTALLATION

        ##*===============================================

        [String]$installPhase = 'Post-Uninstallation'

 

        ## <Perform Post-Uninstallation tasks here>

 

    }

    ElseIf ($deploymentType -ieq 'Repair') {

        ##*===============================================

        ##* PRE-REPAIR

        ##*===============================================

        [String]$installPhase = 'Pre-Repair'

 

 

        ## <Perform Pre-Repair tasks here>

 

        ##*===============================================

        ##* REPAIR

        ##*===============================================

        [String]$installPhase = 'Repair'

 

        ## <Perform Repair tasks here>

      

        # Execute-MSI -Action 'Repair' -Path "" -Parameters "REINSTALL=ALL /qn" -LogName ""

 

'@ ) +"`r`n"+
(

@'

        ##*===============================================

        ##* POST-REPAIR

        ##*===============================================

        [String]$installPhase = 'Post-Repair'

 

        ## <Perform Post-Repair tasks here>

 

    }

    ##*===============================================

    ##* END SCRIPT BODY

    ##*===============================================

 

    ## Call the Exit-Script function to perform final cleanup operations

    Exit-Script -ExitCode $mainExitCode

}

 

Catch {

    [Int32]$mainExitCode = 60001

    [String]$mainErrorMessage = $("Resolve"+"-"+"Error")

    Write-Log -Message $mainErrorMessage -Severity 3 -Source $deployAppScriptFriendlyName

    Show-DialogBox -Text $mainErrorMessage -Icon 'Stop'

    Exit-Script -ExitCode $mainExitCode

}            

'@
)
)

$PSADTcode | Out-File "$PSScriptRoot\Toolkit\Deploy-Application.ps1"
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
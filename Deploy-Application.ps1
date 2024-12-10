<#
.SYNOPSIS

PSApppDeployToolkit - This script performs the installation or uninstallation of an application(s).

.DESCRIPTION

- The script is provided as a template to perform an install or uninstall of an application(s).
- The script either performs an "Install" deployment type or an "Uninstall" deployment type.
- The install deployment type is broken down into 3 main sections/phases: Pre-Install, Install, and Post-Install.

The script dot-sources the AppDeployToolkitMain.ps1 script which contains the logic and functions required to install or uninstall an application.

PSApppDeployToolkit is licensed under the GNU LGPLv3 License - (C) 2023 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham and Muhammad Mashwani).

This program is free software: you can redistribute it and/or modify it under the terms of the GNU Lesser General Public License as published by the
Free Software Foundation, either version 3 of the License, or any later version. This program is distributed in the hope that it will be useful, but
WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License
for more details. You should have received a copy of the GNU Lesser General Public License along with this program. If not, see <http://www.gnu.org/licenses/>.

.PARAMETER DeploymentType

The type of deployment to perform. Default is: Install.

.PARAMETER DeployMode

Specifies whether the installation should be run in Interactive, Silent, or NonInteractive mode. Default is: Interactive. Options: Interactive = Shows dialogs, Silent = No dialogs, NonInteractive = Very silent, i.e. no blocking apps. NonInteractive mode is automatically set if it is detected that the process is not user interactive.

.PARAMETER AllowRebootPassThru

Allows the 3010 return code (requires restart) to be passed back to the parent process (e.g. SCCM) if detected from an installation. If 3010 is passed back to SCCM, a reboot prompt will be triggered.

.PARAMETER TerminalServerMode

Changes to "user install mode" and back to "user execute mode" for installing/uninstalling applications for Remote Desktop Session Hosts/Citrix servers.

.PARAMETER DisableLogging

Disables logging to file for the script. Default is: $false.

.EXAMPLE

powershell.exe -Command "& { & '.\Deploy-Application.ps1' -DeployMode 'Silent'; Exit $LastExitCode }"

.EXAMPLE

powershell.exe -Command "& { & '.\Deploy-Application.ps1' -AllowRebootPassThru; Exit $LastExitCode }"

.EXAMPLE

powershell.exe -Command "& { & '.\Deploy-Application.ps1' -DeploymentType 'Uninstall'; Exit $LastExitCode }"

.EXAMPLE

Deploy-Application.exe -DeploymentType "Install" -DeployMode "Silent"

.INPUTS

None

You cannot pipe objects to this script.

.OUTPUTS

None

This script does not generate any output.

.NOTES

Toolkit Exit Code Ranges:
- 60000 - 68999: Reserved for built-in exit codes in Deploy-Application.ps1, Deploy-Application.exe, and AppDeployToolkitMain.ps1
- 69000 - 69999: Recommended for user customized exit codes in Deploy-Application.ps1
- 70000 - 79999: Recommended for user customized exit codes in AppDeployToolkitExtensions.ps1

.LINK

https://psappdeploytoolkit.com
#>


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
    [switch]$DisableLogging = $false,
    [Parameter(Mandatory = $false)]
    [string]$ProductType = "0"
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
    [String]$appVendor = 'Microsoft'

    ### change index to install different products, indexing begins at 0
    $reqd_config_index = $ProductType
    $installTypeJson = '[{"name":"365 Suite (user, monthly)","installxml":"365suite-user-outlook-monthly.xml","removexml":"remove-365suite.xml","regdetect":"*Microsoft 365*"},
                         {"name":"365 Suite (user, current)","installxml":"365suite-user-outlook-current.xml","removexml":"remove-365suite.xml","regdetect":"*Microsoft 365*"},
                         {"name":"365 Suite (shared, monthly)","installxml":"365suite-shared-outlook-monthly.xml","removexml":"remove-365suite.xml","regdetect":"*Microsoft 365*"},
                         {"name":"365 Suite (shared, current)","installxml":"365suite-shared-outlook-current.xml","removexml":"remove-365suite.xml","regdetect":"*Microsoft 365*"},
                         {"name":"365 Suite - no outlook (shared, monthly)","installxml":"365suite-shared-nooutlook-monthly.xml","removexml":"remove-365suite.xml","regdetect":"*Microsoft 365*"},
                         {"name":"365 Suite - no outlook (shared, current)","installxml":"365suite-shared-nooutlook-current.xml","removexml":"remove-365suite.xml","regdetect":"*Microsoft 365*"},
                         {"name":"Project (user, monthly)","installxml":"project-user-monthly.xml","removexml":"remove-project.xml","regdetect":"*Microsoft Project*"},
                         {"name":"Project (user, current)","installxml":"project-user-current.xml","removexml":"remove-project.xml","regdetect":"*Microsoft Project*"},
                         {"name":"Project (shared, monthly)","installxml":"project-shared-monthly.xml","removexml":"remove-project.xml","regdetect":"*Microsoft Project*"},
                         {"name":"Project (shared, current)","installxml":"project-shared-current.xml","removexml":"remove-project.xml","regdetect":"*Microsoft Project*"},
                         {"name":"Visio (user, monthly)","installxml":"visio-user-monthly.xml","removexml":"remove-visio.xml","regdetect":"*Microsoft Visio*"},
                         {"name":"Visio (user, current)","installxml":"visio-user-current.xml","removexml":"remove-visio.xml","regdetect":"*Microsoft Visio*"},
                         {"name":"Visio (shared, monthly)","installxml":"visio-shared-monthly.xml","removexml":"remove-visio.xml","regdetect":"*Microsoft Visio*"},
                         {"name":"Visio (shared, current)","installxml":"visio-shared-current.xml","removexml":"remove-visio.xml","regdetect":"*Microsoft Visio*"}
                         ]'

    $installTypeConfig = $installTypeJson | ConvertFrom-Json -ErrorAction Stop

    [String]$appName = $installtypeconfig[$reqd_config_index].name

    [String]$appVersion = '2.0'
    [String]$appArch = 'x64'
    [String]$appLang = 'EN'
    [String]$appRevision = '01'
    [String]$appScriptVersion = '1.0.0'
    [String]$appScriptDate = '10/12/2024'
    [String]$appScriptAuthor = ''
    ##*===============================================
    ## Variables: Install Titles (Only set here to override defaults set by the toolkit)
    [String]$installName = ''
    [String]$installTitle = ''

    ##* Do not modify section below
    #region DoNotModify

    ## Variables: Exit Code
    [Int32]$mainExitCode = 0

    ## Variables: Script
    [String]$deployAppScriptFriendlyName = 'Deploy Application'
    [Version]$deployAppScriptVersion = [Version]'3.9.2'
    [String]$deployAppScriptDate = '02/02/2023'
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

        ## Show Welcome Message, close Internet Explorer if required, allow up to 3 deferrals, verify there is enough disk space to complete the install, and persist the prompt
        # Show-InstallationWelcome -CloseApps 'iexplore' -AllowDefer -DeferTimes 3 -CheckDiskSpace -PersistPrompt
        Show-InstallationWelcome -CloseApps 'excel,groove,onenote,outlook,mspub,powerpnt,winword,visio,winproj' -CheckDiskSpace -PersistPrompt

        ## Show Progress Message (with the default message)
        Show-InstallationProgress

        ## <Perform Pre-Installation tasks here>


        ##*===============================================
        ##* INSTALLATION
        ##*===============================================
        [String]$installPhase = 'Installation'

        ## Handle Zero-Config MSI Installations
        If ($useDefaultMsi) {
            [Hashtable]$ExecuteDefaultMSISplat = @{ Action = 'Install'; Path = $defaultMsiFile }; If ($defaultMstFile) {
                $ExecuteDefaultMSISplat.Add('Transform', $defaultMstFile)
            }
            Execute-MSI @ExecuteDefaultMSISplat; If ($defaultMspFiles) {
                $defaultMspFiles | ForEach-Object { Execute-MSI -Action 'Patch' -Path $_ }
            }
        }

        ## <Perform Installation tasks here>

        function Get-ODTURL {
            [String]$MSWebPage = Invoke-RestMethod 'https://www.microsoft.com/en-gb/download/details.aspx?id=49117'
          
            $MSWebPage | ForEach-Object {
                if ($_ -match '"url":"(https://[a-zA-Z0-9-_./]*officedeploymenttool[a-zA-Z0-9-_.]*\.exe)"') {
                    $matches[1]
                }
            }
          
        }

        $OfficeInstallDownloadPath = 'C:\temp\Office365Install'

        $VerbosePreference = 'Continue'
        $ErrorActionPreference = 'Stop'
        $CleanUpInstallFiles = $true


        if (-Not(Test-Path "$OfficeInstallDownloadPath" )) {
            New-Folder -Path $OfficeInstallDownloadPath
        }

        $ConfigurationXMLFile = "$dirSupportFiles\$($installtypeconfig[$reqd_config_index].installxml)"
        Write-Log -Message "Using xml configuration $ConfigurationXMLFile for deplyment type $appName"

        $ODTInstallLink = Get-ODTURL

        #Download the Office Deployment Tool
        Write-Log -Message 'Downloading the Office Deployment Tool...'
        try {
            Invoke-WebRequest -Uri $ODTInstallLink -OutFile "$OfficeInstallDownloadPath\ODTSetup.exe"
        }
        catch {
            Write-Log -Severity 3 -Message 'There was an error downloading the Office Deployment Tool.'
            Write-Log -Severity 3 -Message 'Please verify the below link is valid:'
            Write-Log -Severity 3 -Message "$ODTInstallLink"
            Exit-Script -ExitCode 1
        }

        #Run the Office Deployment Tool setup
        Write-Log -Message 'Running the Office Deployment Tool...'
        try {
            Execute-Process -Path "$OfficeInstallDownloadPath\ODTSetup.exe" -Parameters "/quiet /extract:$OfficeInstallDownloadPath"
        }
        catch {
            Write-Log -Severity 3 -Message 'Error running the Office Deployment Tool. The error is below:'
            Write-Log -Severity 3 -Message "$_"
            Exit-Script -ExitCode 1
        }

        #Run the Microsoft 365 Apps install
        Write-Log -Message 'Downloading and installing Microsoft 365'
        try {
            Execute-Process -Path "$OfficeInstallDownloadPath\Setup.exe" -Parameters "/configure $ConfigurationXMLFile"
        }
        catch {
            Write-Log -Severity 3 -Message 'Error running the Office install. The error is below:'
            Write-Log -Severity 3 -Message "$_"
            Exit-Script -ExitCode 1
        }
  

        #Check if Microsoft 365 suite was installed correctly.
        $RegLocations = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
                          'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
        )

        $OfficeInstalled = $False
        $regdetect = $installtypeconfig[$reqd_config_index].regdetect
        foreach ($Key in (Get-ChildItem $RegLocations) ) {
            #if ($Key.GetValue('DisplayName') -like '*Microsoft 365*') {
            if ($Key.GetValue('DisplayName') -like "$regdetect") {
                $OfficeVersionInstalled = $Key.GetValue('DisplayName')
                $OfficeInstalled = $True
            }
        }

        if ($OfficeInstalled) {
            Write-Log -Message "$($OfficeVersionInstalled) installed successfully!"
        } else {
            Write-Log -Severity 2 -Message "$appName was not detected after the install ran"
        }

        # sleep for a couple of minutes to let the installation finish
        Start-Sleep -Seconds 120

        if ($CleanUpInstallFiles) {
            Remove-Folder -Path "$OfficeInstallDownloadPath" -ErrorAction SilentlyContinue
        }


        ##*===============================================
        ##* POST-INSTALLATION
        ##*===============================================
        [String]$installPhase = 'Post-Installation'

        ## <Perform Post-Installation tasks here>
		Set-RegistryKey -Key 'HKEY_LOCAL_MACHINE\SOFTWARE\Intune_Installations' -Name "$appVendor $appName $appVersion" -Value '"Installed"'-Type String

        ## Display a message at the end of the install
        If (-not $useDefaultMsi) { }
    }
    ElseIf ($deploymentType -ieq 'Uninstall') {
        ##*===============================================
        ##* PRE-UNINSTALLATION
        ##*===============================================
        [String]$installPhase = 'Pre-Uninstallation'

        ## Show Welcome Message, close Internet Explorer with a 60 second countdown before automatically closing
        Show-InstallationWelcome -CloseApps 'excel,groove,onenote,outlook,mspub,powerpnt,winword,visio,winproj' -CloseAppsCountdown 60

        ## Show Progress Message (with the default message)
        Show-InstallationProgress

        ## <Perform Pre-Uninstallation tasks here>


        ##*===============================================
        ##* UNINSTALLATION
        ##*===============================================
        [String]$installPhase = 'Uninstallation'

        ## Handle Zero-Config MSI Uninstallations
        If ($useDefaultMsi) {
            [Hashtable]$ExecuteDefaultMSISplat = @{ Action = 'Uninstall'; Path = $defaultMsiFile }; If ($defaultMstFile) {
                $ExecuteDefaultMSISplat.Add('Transform', $defaultMstFile)
            }
            Execute-MSI @ExecuteDefaultMSISplat
        }

        function Get-ODTURL {
            [String]$MSWebPage = Invoke-RestMethod 'https://www.microsoft.com/en-us/download/confirmation.aspx?id=49117'
          
            $MSWebPage | ForEach-Object {
                if ($_ -match 'url=(https://.*officedeploymenttool.*\.exe)') {
                    $matches[1]
                }
            }
          
        }

        $OfficeInstallDownloadPath = 'C:\temp\Office365Install'

        $VerbosePreference = 'Continue'
        $ErrorActionPreference = 'Stop'
        $CleanUpInstallFiles = $true


        if (-Not(Test-Path "$OfficeInstallDownloadPath" )) {
            New-Folder -Path $OfficeInstallDownloadPath
        }

        $ConfigurationXMLFile = "$dirSupportFiles\$($installtypeconfig[$reqd_config_index].removexml)"
        Write-Log -Message "Using xml configuration $ConfigurationXMLFile for deplyment type $appName"

        $ODTInstallLink = Get-ODTURL

        #Download the Office Deployment Tool
        Write-Log -Message 'Downloading the Office Deployment Tool...'
        try {
            Invoke-WebRequest -Uri $ODTInstallLink -OutFile "$OfficeInstallDownloadPath\ODTSetup.exe"
        }
        catch {
            Write-Log -Severity 3 -Message 'There was an error downloading the Office Deployment Tool.'
            Write-Log -Severity 3 -Message 'Please verify the below link is valid:'
            Write-Log -Severity 3 -Message "$ODTInstallLink"
            Exit-Script -ExitCode 1
        }

        #Run the Office Deployment Tool setup
        Write-Log -Message 'Running the Office Deployment Tool...'
        try {
            Execute-Process -Path "$OfficeInstallDownloadPath\ODTSetup.exe" -Parameters "/quiet /extract:$OfficeInstallDownloadPath"
        }
        catch {
            Write-Log -Severity 3 -Message 'Error running the Office Deployment Tool. The error is below:'
            Write-Log -Severity 3 -Message "$_"
            Exit-Script -ExitCode 1
        }

        #Run the Microsoft 365 Apps install
        Write-Log -Message 'Removing Microsoft 365'
        try {
            Execute-Process -Path "$OfficeInstallDownloadPath\Setup.exe" -Parameters "/configure $ConfigurationXMLFile"
        }
        catch {
            Write-Log -Severity 3 -Message 'Error running the Office install. The error is below:'
            Write-Log -Severity 3 -Message "$_"
            Exit-Script -ExitCode 1
        }
  
        #Check if Microsoft 365 suite was uninstalled correctly.
        $RegLocations = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
                          'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
        )

        $OfficeInstalled = $False
        $regdetect = $installtypeconfig[$reqd_config_index].regdetect
        foreach ($Key in (Get-ChildItem $RegLocations) ) {
            if ($Key.GetValue('DisplayName') -like "$regdetect") {
                $OfficeVersionInstalled = $Key.GetValue('DisplayName')
                $OfficeInstalled = $True
            }
        }

        if ($OfficeInstalled) {
            Write-Log -Severity 2 -Message "$($OfficeVersionInstalled) was detected, uninstall was unsuccessful"
        } else {
            Write-Log -Message "$appName was not detected after the uninstall ran, successful"
        }

        if ($CleanUpInstallFiles) {
            Remove-Folder -Path "$OfficeInstallDownloadPath" -ErrorAction SilentlyContinue
        }

        ## <Perform Uninstallation tasks here>
		Remove-RegistryKey -Key 'HKEY_LOCAL_MACHINE\SOFTWARE\Intune_Installations' -Name "$appVendor $appName $appVersion"

        ##*===============================================
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

        ## Show Welcome Message, close Internet Explorer with a 60 second countdown before automatically closing
        Show-InstallationWelcome -CloseApps 'iexplore' -CloseAppsCountdown 60

        ## Show Progress Message (with the default message)
        Show-InstallationProgress

        ## <Perform Pre-Repair tasks here>

        ##*===============================================
        ##* REPAIR
        ##*===============================================
        [String]$installPhase = 'Repair'

        ## Handle Zero-Config MSI Repairs
        If ($useDefaultMsi) {
            [Hashtable]$ExecuteDefaultMSISplat = @{ Action = 'Repair'; Path = $defaultMsiFile; }; If ($defaultMstFile) {
                $ExecuteDefaultMSISplat.Add('Transform', $defaultMstFile)
            }
            Execute-MSI @ExecuteDefaultMSISplat
        }
        ## <Perform Repair tasks here>

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
    [String]$mainErrorMessage = "$(Resolve-Error)"
    Write-Log -Message $mainErrorMessage -Severity 3 -Source $deployAppScriptFriendlyName
    Show-DialogBox -Text $mainErrorMessage -Icon 'Stop'
    Exit-Script -ExitCode $mainExitCode
}

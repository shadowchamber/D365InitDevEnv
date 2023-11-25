<# Common PowerShell functions used by the Dynamics AX 7.0 build process. #>

<#
.SYNOPSIS
    Import this module to get a Write-Message function for controlling output.
    
.DESCRIPTION
    This script can be imported to log messages of type message, error or warning. 
    If no $LogPath variable is defined it will write to host using either Write-Host,
    Write-Warning, Write-Error, or Write-Verbose.

.NOTES
    When running through automation, set the $LogPath variable to redirect
    all output to a log file rather than the console. Can be set in calling script.

    Copyright © 2016 Microsoft. All rights reserved.
#>
function Write-Message
{
    [Cmdletbinding()]
    Param([string]$Message, [switch]$Error, [switch]$Warning, [switch]$Diag, [string]$LogPath = $PSCmdlet.GetVariableValue("LogPath"))

    # Get verbose preference from caller.
    $VerbosePreference = $PSCmdlet.GetVariableValue("VerbosePreference")

    if ($LogPath)
    {
        # For log files use full UTC time stamp.
        "$([DateTime]::UtcNow.ToString("s")): $($Message)" | Out-File -FilePath $LogPath -Append
    }
    else
    {
        # For writing to host use a local time stamp.
        [string]$FormattedMessage = "$([DateTime]::Now.ToLongTimeString()): $($Message)"
        
        # If message is of type Error, use Write-Error.
        if ($Error)
        {
            Write-Error $FormattedMessage
        }
        else
        {
            # If message is of type Warning, use Write-Warning.
            if ($Warning)
            {
                Write-Warning $FormattedMessage
            }
            else
            {
                # If message is of type Verbose, use Write-Verbose.
                if ($Diag)
                {
                    Write-Verbose $FormattedMessage
                }
                else
                {
                    Write-Host $FormattedMessage
                }
            }
        }
    }
}

<#
.SYNOPSIS
    Set DynamicsSDK machine wide environment variables. These are used by the
    build process.
#>
function Set-AX7SdkEnvironmentVariables
{
    [Cmdletbinding()]
    Param([string]$DynamicsSDK)

    # Get verbose preference from caller.
    $VerbosePreference = $PSCmdlet.GetVariableValue("VerbosePreference")

    # If specified, save DynamicsSDK in registry and machine wide environment variable.
    if ($DynamicsSDK)
    {
        Write-Message "- Setting machine wide DynamicsSDK environment variable: $DynamicsSDK" -Diag
        [Environment]::SetEnvironmentVariable("DynamicsSDK", $DynamicsSDK, "Machine")
    }
    else
    {
        Write-Message "- No DynamicsSDK value specified. No environment variable will be set." -Diag
    }
}

<#
.SYNOPSIS
    Set DynamicsSDK, TeamFoundationServerUrl, AosWebsiteName and BackupPath registry values.
    These are used by the build process.
#>
function Set-AX7SdkRegistryValues
{
    [Cmdletbinding()]
    Param([string]$DynamicsSDK, [string]$TeamFoundationServerUrl, [string]$AosWebsiteName, [string]$BackupPath)

    # Get verbose preference from caller.
    $VerbosePreference = $PSCmdlet.GetVariableValue("VerbosePreference")

    $RegPath = "HKLM:\SOFTWARE\Microsoft\Dynamics\AX\7.0\SDK"

    if (!(Test-Path -Path $RegPath))
    {
        Write-Message "- Creating new Dynamics SDK registry key: $RegPath" -Diag
        $RegKey = New-Item -Path $RegPath -Force
    }

    # If specified, save DynamicsSDK in registry.
    if ($DynamicsSDK)
    {
        Write-Message "- Setting DynamicsSDK registry value: $DynamicsSDK" -Diag
        $RegValue = New-ItemProperty -Path $RegPath -Name "DynamicsSDK" -Value $DynamicsSDK -Force
    }
    else
    {
        Write-Message "- No DynamicsSDK value specified. No registry value will be set" -Diag
    }
    
    # If specified, save TeamFoundationServerUrl in registry.
    if ($TeamFoundationServerUrl)
    {
        Write-Message "- Setting TeamFoundationServerUrl registry value: $TeamFoundationServerUrl" -Diag
        $RegValue = New-ItemProperty -Path $RegPath -Name "TeamFoundationServerUrl" -Value $TeamFoundationServerUrl -Force
    }
    else
    {
        Write-Message "- No TeamFoundationServerUrl value specified. No registry value will be set." -Diag
    }

    # If specified, save AosWebsiteName in registry.
    if ($AosWebsiteName)
    {
        Write-Message "- Setting AosWebsiteName registry value: $AosWebsiteName" -Diag
        $RegValue = New-ItemProperty -Path $RegPath -Name "AosWebsiteName" -Value $AosWebsiteName -Force
    }
    else
    {
        Write-Message "- No AosWebsiteName value specified. No registry value will be set." -Diag
    }

    # If specified, save BackupPath in registry.
    if ($BackupPath)
    {
        Write-Message "- Setting BackupPath registry value: $BackupPath" -Diag
        $RegValue = New-ItemProperty -Path $RegPath -Name "BackupPath" -Value $BackupPath -Force
    }
    else
    {
        Write-Message "- No BackupPath value specified. No registry value will be set." -Diag
    }
}

<#
.SYNOPSIS
    Set values in the Dynamics SDK registry key from the AOS web config.
    These are used by the build process and read in properties of project and
    target files.
#>
function Set-AX7SdkRegistryValuesFromAosWebConfig
{
    [Cmdletbinding()]
    Param([string]$AosWebsiteName)
    
    # Get verbose preference from caller.
    $VerbosePreference = $PSCmdlet.GetVariableValue("VerbosePreference")

    $RegPath = "HKLM:\SOFTWARE\Microsoft\Dynamics\AX\7.0\SDK"

    if (!(Test-Path -Path $RegPath))
    {
        Write-Message "- Creating new Dynamics SDK registry key: $RegPath" -Diag
        $RegKey = New-Item -Path $RegPath -Force
    }
    
    # If specified, save AosWebsiteName in registry.
    if ($AosWebsiteName)
    {
        Write-Message "- Setting AosWebsiteName registry value: $AosWebsiteName" -Diag
        $RegValue = New-ItemProperty -Path $RegPath -Name "AosWebsiteName" -Value $AosWebsiteName -Force
    }
    else
    {
        $RegKey = Get-ItemProperty -Path $RegPath
        $AosWebsiteName = $RegKey.AosWebsiteName
        if ($AosWebsiteName)
        {
            Write-Message "- No AOS website name specified. Using existing value from registry: $AosWebsiteName" -Diag
        }
        else
        {
            throw "No AOS website name specified and no existing value found in registry at: $RegPath"
        }
    }

    # Get AOS web.config and extract values to save in Dynamics SDK registry. These will be read
    # by the MSBuild projects.
    $WebConfigPath = Get-AX7DeploymentAosWebConfigPath -WebsiteName $AosWebsiteName
    
    if ($WebConfigPath -and (Test-Path -Path $WebConfigPath -PathType Leaf))
    {
        $BinariesPath = Get-AX7DeploymentBinariesPath -WebConfigPath $WebConfigPath
        if ($BinariesPath)
        {
            Write-Message "- Setting BinariesPath registry value: $BinariesPath" -Diag
            $RegValue = New-ItemProperty -Path $RegPath -Name "BinariesPath" -Value $BinariesPath -Force
        }
        else
        {
            Write-Message "- No BinariesPath could be found in AOS web.config: $WebConfigPath" -Warning
        }
    
        $MetadataPath = Get-AX7DeploymentMetadataPath -WebConfigPath $WebConfigPath
        if ($MetadataPath)
        {
            Write-Message "- Setting MetadataPath registry value: $MetadataPath" -Diag
            $RegValue = New-ItemProperty -Path $RegPath -Name "MetadataPath" -Value $MetadataPath -Force
        }
        else
        {
            Write-Message "- No MetadataPath could be found in AOS web.config: $WebConfigPath" -Warning
        }
    
        $PackagesPath = Get-AX7DeploymentPackagesPath -WebConfigPath $WebConfigPath
        if ($PackagesPath)
        {
            Write-Message "- Setting PackagesPath registry value: $PackagesPath" -Diag
            $RegValue = New-ItemProperty -Path $RegPath -Name "PackagesPath" -Value $PackagesPath -Force
        }
        else
        {
            Write-Message "- No PackagesPath could be found in AOS web.config: $WebConfigPath" -Warning
        }

        $DatabaseName = Get-AX7DeploymentDatabaseName -WebConfigPath $WebConfigPath
        if ($DatabaseName)
        {
            Write-Message "- Setting DatabaseName registry value: $DatabaseName" -Diag
            $RegValue = New-ItemProperty -Path $RegPath -Name "DatabaseName" -Value $DatabaseName -Force
        }
        else
        {
            Write-Message "- No DatabaseName could be found in AOS web.config: $WebConfigPath" -Warning
        }

        $DatabaseServer = Get-AX7DeploymentDatabaseServer -WebConfigPath $WebConfigPath
        if ($DatabaseServer)
        {
            Write-Message "- Setting DatabaseServer registry value: $DatabaseServer" -Diag
            $RegValue = New-ItemProperty -Path $RegPath -Name "DatabaseServer" -Value $DatabaseServer -Force
        }
        else
        {
            Write-Message "- No DatabaseServer could be found in AOS web.config: $WebConfigPath" -Warning
        }
    }
    else
    {
        throw "No AOS web config could be found for AOS website name: $AosWebsiteName"
    }
}

<#
.SYNOPSIS
    Get the Dynamics SDK path from the Dynamics SDK registry key.

.NOTES
    Throws exception if the Dynamics SDK registry path is not found.

.OUTPUTS
    System.String. The full path to the Dynamics SDK files.
#>
function Get-AX7SdkPath
{
    [Cmdletbinding()]
    Param()

    # Get verbose preference from caller.
    $VerbosePreference = $PSCmdlet.GetVariableValue("VerbosePreference")

    [string]$DynamicsSdk = $null

    $RegPath = "HKLM:\SOFTWARE\Microsoft\Dynamics\AX\7.0\SDK"
    
    # Get the Dynamics SDK registry key (throws if not found).
    Write-Message "- Getting Dynamics SDK registry key..." -Diag
    $RegKey = Get-ItemProperty -Path $RegPath

    if ($RegKey -ne $null)
    {
        $DynamicsSdk = $RegKey.DynamicsSDK
        if ($DynamicsSdk)
        {
            Write-Message "- Found Dynamics SDK path: $DynamicsSdk" -Diag
        }
        else
        {
            Write-Message "- No Dynamics SDK path found in registry." -Diag
        }
    }

    return $DynamicsSdk
}

<#
.SYNOPSIS
    Get the Dynamics SDK backup path from the Dynamics SDK registry key.

.NOTES
    Throws exception if the Dynamics SDK registry path is not found.

.OUTPUTS
    System.String. The full path to the backup path.
#>
function Get-AX7SdkBackupPath
{
    [Cmdletbinding()]
    Param()

    # Get verbose preference from caller.
    $VerbosePreference = $PSCmdlet.GetVariableValue("VerbosePreference")

    [string]$BackupPath = $null

    $RegPath = "HKLM:\SOFTWARE\Microsoft\Dynamics\AX\7.0\SDK"
    
    # Get the Dynamics SDK registry key (throws if not found).
    Write-Message "- Getting Dynamics SDK registry key..." -Diag
    $RegKey = Get-ItemProperty -Path $RegPath

    if ($RegKey -ne $null)
    {
        $BackupPath = $RegKey.BackupPath
        if ($BackupPath)
        {
            Write-Message "- Found backup path: $BackupPath" -Diag
        }
        else
        {
            Write-Message "- No backup path found in registry." -Diag
        }
    }

    return $BackupPath
}

<#
.SYNOPSIS
    Get the VSO/Team Foundation Server URL from the Dynamics SDK registry key.

.NOTES
    Throws exception if the Dynamics SDK registry path is not found.

.OUTPUTS
    System.String. The VSO/Team Foundation Server URL.
#>
function Get-AX7SdkTeamFoundationServerUrl
{
    [Cmdletbinding()]
    Param()

    # Get verbose preference from caller.
    $VerbosePreference = $PSCmdlet.GetVariableValue("VerbosePreference")

    [string]$TeamFoundationServerUrl = $null

    $RegPath = "HKLM:\SOFTWARE\Microsoft\Dynamics\AX\7.0\SDK"
    
    # Get the Dynamics SDK registry key (throws if not found).
    Write-Message "- Getting Dynamics SDK registry key..." -Diag
    $RegKey = Get-ItemProperty -Path $RegPath

    if ($RegKey -ne $null)
    {
        $TeamFoundationServerUrl = $RegKey.TeamFoundationServerUrl
        if ($TeamFoundationServerUrl)
        {
            Write-Message "- Found Team Foundation Server URL: $TeamFoundationServerUrl" -Diag
        }
        else
        {
            Write-Message "- No Team Foundation Server URL found in registry." -Diag
        }
    }

    return $TeamFoundationServerUrl
}

<#
.SYNOPSIS
    Get the Dynamics AX deployment database name from the Dynamics SDK registry key.

.NOTES
    Throws exception if the Dynamics SDK registry path is not found.

.OUTPUTS
    System.String. The database name.
#>
function Get-AX7SdkDeploymentDatabaseName
{
    [Cmdletbinding()]
    Param()

    # Get verbose preference from caller.
    $VerbosePreference = $PSCmdlet.GetVariableValue("VerbosePreference")

    [string]$DatabaseName = $null

    $RegPath = "HKLM:\SOFTWARE\Microsoft\Dynamics\AX\7.0\SDK"
    
    # Get the Dynamics SDK registry key (throws if not found).
    Write-Message "- Getting Dynamics SDK registry key..." -Diag
    $RegKey = Get-ItemProperty -Path $RegPath

    if ($RegKey -ne $null)
    {
        $DatabaseName = $RegKey.DatabaseName
        if ($DatabaseName)
        {
            Write-Message "- Found database name: $DatabaseName" -Diag
        }
        else
        {
            Write-Message "- No database name found in registry." -Diag
        }
    }

    return $DatabaseName
}

<#
.SYNOPSIS
    Get the Dynamics AX deployment database server from the Dynamics SDK registry key.

.NOTES
    Throws exception if the Dynamics SDK registry path is not found.

.OUTPUTS
    System.String. The database server name.
#>
function Get-AX7SdkDeploymentDatabaseServer
{
    [Cmdletbinding()]
    Param()

    # Get verbose preference from caller.
    $VerbosePreference = $PSCmdlet.GetVariableValue("VerbosePreference")

    [string]$DatabaseServer = $null

    $RegPath = "HKLM:\SOFTWARE\Microsoft\Dynamics\AX\7.0\SDK"
    
    # Get the Dynamics SDK registry key (throws if not found).
    Write-Message "- Getting Dynamics SDK registry key..." -Diag
    $RegKey = Get-ItemProperty -Path $RegPath

    if ($RegKey -ne $null)
    {
        $DatabaseServer = $RegKey.DatabaseServer
        if ($DatabaseServer)
        {
            Write-Message "- Found database server: $DatabaseServer" -Diag
        }
        else
        {
            Write-Message "- No database server found in registry." -Diag
        }
    }

    return $DatabaseServer
}

<#
.SYNOPSIS
    Get the Dynamics AX deployment AOS website name from the Dynamics SDK registry key.

.NOTES
    Throws exception if the Dynamics SDK registry path is not found.

.OUTPUTS
    System.String. The deployment AOS website name.
#>
function Get-AX7SdkDeploymentAosWebsiteName
{
    [Cmdletbinding()]
    Param()
    
    # Get verbose preference from caller.
    $VerbosePreference = $PSCmdlet.GetVariableValue("VerbosePreference")

    [string]$AosWebsiteName = $null

    $RegPath = "HKLM:\SOFTWARE\Microsoft\Dynamics\AX\7.0\SDK"
    
    # Get the Dynamics SDK registry key (throws if not found).
    Write-Message "- Getting Dynamics SDK registry key..." -Diag
    $RegKey = Get-ItemProperty -Path $RegPath

    if ($RegKey -ne $null)
    {
        Write-Message "- Getting deployment AOS website name..." -Diag
        $AosWebsiteName = $RegKey.AosWebsiteName
        if ($AosWebsiteName)
        {
            Write-Message "- Found deployment AOS website name: $AosWebsiteName" -Diag
        }
        else
        {
            Write-Message "- No deployment AOS website name found in registry." -Diag
        }
    }

    return $AosWebsiteName
}

<#
.SYNOPSIS
    Get the Dynamics AX deployment binaries path from Dynamics SDK registry key.

.NOTES
    Throws exception if the Dynamics SDK registry path is not found.

.OUTPUTS
    System.String. The full path to the deployment binaries.
#>
function Get-AX7SdkDeploymentBinariesPath
{
    [Cmdletbinding()]
    Param()

    # Get verbose preference from caller.
    $VerbosePreference = $PSCmdlet.GetVariableValue("VerbosePreference")

    [string]$BinariesPath = $null

    $RegPath = "HKLM:\SOFTWARE\Microsoft\Dynamics\AX\7.0\SDK"
    
    # Get the Dynamics SDK registry key (throws if not found).
    Write-Message "- Getting Dynamics SDK registry key..." -Diag
    $RegKey = Get-ItemProperty -Path $RegPath

    if ($RegKey -ne $null)
    {
        Write-Message "- Getting deployment binaries path..." -Diag
        $BinariesPath = $RegKey.BinariesPath
        if ($BinariesPath)
        {
            Write-Message "- Found deployment binaries path: $BinariesPath" -Diag
        }
        else
        {
            Write-Message "- No deployment binaries path found in registry." -Diag
        }
    }

    return $BinariesPath
}

<#
.SYNOPSIS
    Get the Dynamics AX deployment metadata path from Dynamics SDK registry key.

.NOTES
    Throws exception if the Dynamics SDK registry path is not found.

.OUTPUTS
    System.String. The full path to the deployment metadata.
#>
function Get-AX7SdkDeploymentMetadataPath
{
    [Cmdletbinding()]
    Param()

    # Get verbose preference from caller.
    $VerbosePreference = $PSCmdlet.GetVariableValue("VerbosePreference")

    [string]$MetadataPath = $null

    $RegPath = "HKLM:\SOFTWARE\Microsoft\Dynamics\AX\7.0\SDK"
    
    # Get the Dynamics SDK registry key (throws if not found).
    Write-Message "- Getting Dynamics SDK registry key..." -Diag
    $RegKey = Get-ItemProperty -Path $RegPath

    if ($RegKey -ne $null)
    {
        Write-Message "- Getting deployment metadata path..." -Diag
        $MetadataPath = $RegKey.MetadataPath
        if ($MetadataPath)
        {
            Write-Message "- Found deployment metadata path: $MetadataPath" -Diag
        }
        else
        {
            Write-Message "- No deployment metadata path found in registry." -Diag
        }
    }

    return $MetadataPath
}

<#
.SYNOPSIS
    Get the Dynamics AX deployment packages path from Dynamics SDK registry key.

.NOTES
    Throws exception if the Dynamics SDK registry path is not found.

.OUTPUTS
    System.String. The full path to the deployment packages.
#>
function Get-AX7SdkDeploymentPackagesPath
{
    [Cmdletbinding()]
    Param()

    # Get verbose preference from caller.
    $VerbosePreference = $PSCmdlet.GetVariableValue("VerbosePreference")

    [string]$PackagesPath = $null

    $RegPath = "HKLM:\SOFTWARE\Microsoft\Dynamics\AX\7.0\SDK"
    
    # Get the Dynamics SDK registry key (throws if not found).
    Write-Message "- Getting Dynamics SDK registry key..." -Diag
    $RegKey = Get-ItemProperty -Path $RegPath

    if ($RegKey -ne $null)
    {
        Write-Message "- Getting deployment packages path..." -Diag
        $PackagesPath = $RegKey.PackagesPath
        if ($PackagesPath)
        {
            Write-Message "- Found deployment packages path: $PackagesPath" -Diag
        }
        else
        {
            Write-Message "- No deployment packages path found in registry." -Diag
        }
    }

    return $PackagesPath
}

<#
.SYNOPSIS
    Stop the IIS service.
#>
function Stop-IIS
{
    [Cmdletbinding()]
    Param()
        
    # Get verbose preference from caller.
    $VerbosePreference = $PSCmdlet.GetVariableValue("VerbosePreference")

    Write-Message "- Calling IISReset /STOP to stop IIS..." -Diag

    $IISResetOutput = & IISReset /STOP
        
    # Check exit code to make sure the service was correctly removed.
    $IISResetExitCode = [int]$LASTEXITCODE
                
    # Log output if any.
    if ($IISResetOutput -and $IISResetOutput.Count -gt 0)
    {
        $IISResetOutput | % { Write-Message $_ -Diag }
    }

    Write-Message "- IISReset completed with exit code: $IISResetExitCode" -Diag
    if ($IISResetExitCode -ne 0)
	{
		throw "IISReset returned an unexpected exit code: $IISResetExitCode"
	}

    Write-Message "- IIS stopped successfully." -Diag
}

<#
.SYNOPSIS
    Start the IIS service.
#>
function Start-IIS
{
    [Cmdletbinding()]
    Param()
    
    # Get verbose preference from caller.
    $VerbosePreference = $PSCmdlet.GetVariableValue("VerbosePreference")

    Write-Message "- Calling IISReset /START to start IIS..." -Diag

    $IISResetOutput = & IISReset /START
        
    # Check exit code to make sure the service was correctly removed.
    $IISResetExitCode = [int]$LASTEXITCODE
                
    # Log output if any.
    if ($IISResetOutput -and $IISResetOutput.Count -gt 0)
    {
        $IISResetOutput | % { Write-Message $_ -Diag }
    }

    Write-Message "- IISReset completed with exit code: $IISResetExitCode" -Diag
    if ($IISResetExitCode -ne 0)
	{
		throw "IISReset returned an unexpected exit code: $IISResetExitCode"
	}

    Write-Message "- IIS started successfully." -Diag
}

<#
.SYNOPSIS
    Restart the IIS service.
#>
function Restart-IIS
{
    [Cmdletbinding()]
    Param()
    
    # Get verbose preference from caller.
    $VerbosePreference = $PSCmdlet.GetVariableValue("VerbosePreference")

    Write-Message "- Calling IISReset /RESTART to restart IIS..." -Diag
    
    $IISResetOutput = & IISReset /RESTART
    
    # Check exit code to make sure the service was correctly removed.
    $IISResetExitCode = [int]$LASTEXITCODE
    
    # Log output if any.
    if ($IISResetOutput -and $IISResetOutput.Count -gt 0)
    {
        $IISResetOutput | % { Write-Message $_ -Diag }
    }

    Write-Message "- IISReset completed with exit code: $IISResetExitCode" -Diag
    if ($IISResetExitCode -ne 0)
	{
		throw "IISReset returned an unexpected exit code: $IISResetExitCode"
	}

    Write-Message "- IIS restarted successfully." -Diag
}

<#
.SYNOPSIS
    Get the Dynamics AX 7.0 deployment AOS website path.

.DESCRIPTION
    If a website name is not specified it will try to use the default website
    name used by the deployment process.
#>
function Get-AX7DeploymentAosWebsite
{
    [Cmdletbinding()]
    Param([string]$WebsiteName)
    
    # Get verbose preference from caller.
    $VerbosePreference = $PSCmdlet.GetVariableValue("VerbosePreference")

    # Import only the functions needed (Too noisy with Verbose).
    Import-Module -Name "WebAdministration" -Function "Get-Website" -Verbose:$false

    [Microsoft.IIs.PowerShell.Framework.ConfigurationElement]$Website = $null

    if ($WebsiteName)
    {
        # Use specified website name.
        $Website = Get-Website -Name $WebsiteName
    }
    else
    {
        # Try default service model website name.
        $Website = Get-Website -Name "AosService"
        if (!$Website)
        {
            # Try default deploy website name.
            $Website = Get-Website -Name "AosWebApplication"
        }
    }

    return $Website
}

<#
.SYNOPSIS
    Get the Dynamics AX 7.0 deployment AOS website path.
#>
function Get-AX7DeploymentAosWebsitePath
{
    [Cmdletbinding()]
    Param([string]$WebsiteName)
    
    # Get verbose preference from caller.
    $VerbosePreference = $PSCmdlet.GetVariableValue("VerbosePreference")

    [string]$AosWebsitePath = $null

    # Get website and its physical path.
    $Website = Get-AX7DeploymentAosWebsite -WebsiteName $WebsiteName
    if ($Website)
    {
        $AosWebsitePath = $Website.physicalPath
    }
    else
    {
        throw "No AOS website could be found in IIS."
    }

    return $AosWebsitePath
}

<#
.SYNOPSIS
    Get the Dynamics AX 7.0 deployment AOS web config path.
#>
function Get-AX7DeploymentAosWebConfigPath
{
    [Cmdletbinding()]
    Param([string]$WebsiteName)
    
    # Get verbose preference from caller.
    $VerbosePreference = $PSCmdlet.GetVariableValue("VerbosePreference")

    [string]$AosWebConfigPath = $null

    [string]$AosWebsitePath = Get-AX7DeploymentAosWebsitePath -WebsiteName $WebsiteName

    if ($AosWebsitePath)
    {
        $AosWebConfigPath = Join-Path -Path $AosWebsitePath -ChildPath "web.config"
    }

    return $AosWebConfigPath
}

<#
.SYNOPSIS
    Get the Dynamics AX 7.0 deployment AOS wif config path.
#>
function Get-AX7DeploymentAosWifConfigPath
{
    [Cmdletbinding()]
    Param([string]$WebsiteName)
    
    # Get verbose preference from caller.
    $VerbosePreference = $PSCmdlet.GetVariableValue("VerbosePreference")
        
    [string]$AosWifConfigPath = $null

    [string]$AosWebsitePath = Get-AX7DeploymentAosWebsitePath -WebsiteName $WebsiteName

    if ($AosWebsitePath)
    {
        $AosWifConfigPath = Join-Path -Path $AosWebsitePath -ChildPath "wif.config"
    }

    return $AosWifConfigPath
}

<#
.SYNOPSIS
    Get the setting value from the specified web.config file path mathing
    the specified setting name.
#>
function Get-AX7DeploymentAosWebConfigSetting
{
    [Cmdletbinding()]
    Param([string]$WebConfigPath, [string]$Name)
    
    # Get verbose preference from caller.
    $VerbosePreference = $PSCmdlet.GetVariableValue("VerbosePreference")

    [string]$SettingValue = $null

    if (Test-Path -Path $WebConfigPath -PathType Leaf)
    {
        [xml]$WebConfig = Get-Content -Path $WebConfigPath
        if ($WebConfig)
        {
            $XPath = "/configuration/appSettings/add[@key='$($Name)']"
            $KeyNode = $WebConfig.SelectSingleNode($XPath)
            if ($KeyNode)
            {
                $SettingValue = $KeyNode.Value
            }
            else
            {
                throw "Failed to find setting in web.config at: $XPath"
            }
        }
        else
        {
            throw "Failed to read web.config content from: $WebConfigPath"
        }
    }
    else
    {
        throw "The specified web.config file could not be found at: $WebConfigPath"
    }

    return $SettingValue
}

<#
.SYNOPSIS
    Get the Dynamics AX 7.0 deployment binaries path.

.DESCRIPTION
    Value is extracted from the specified web.config file path of the default
    AOS web config if no web config path is specified.
#>
function Get-AX7DeploymentBinariesPath
{
    [Cmdletbinding()]
    Param([string]$WebConfigPath)
    
    # Get verbose preference from caller.
    $VerbosePreference = $PSCmdlet.GetVariableValue("VerbosePreference")

    [string]$BinariesPath = $null

    if (!$WebConfigPath)
    {
        $WebConfigPath = Get-AX7DeploymentAosWebConfigPath
    }
    # TODO: Correct this if Common.BinDir will ever be fixed to contain Bin.
    $BinariesPath = Get-AX7DeploymentAosWebConfigSetting -WebConfigPath $WebConfigPath -Name "Common.BinDir"
    if (!($BinariesPath -imatch "\\Bin$"))
    {
        $BinariesPath = Join-Path -Path $BinariesPath -ChildPath "Bin"
    }    

    return $BinariesPath
}

<#
.SYNOPSIS
    Get the Dynamics AX 7.0 deployment metadata path.

.DESCRIPTION
    Value is extracted from the specified web.config file path of the default
    AOS web config if no web config path is specified.
#>
function Get-AX7DeploymentMetadataPath
{
    [Cmdletbinding()]
    Param([string]$WebConfigPath)
    
    # Get verbose preference from caller.
    $VerbosePreference = $PSCmdlet.GetVariableValue("VerbosePreference")
        
    [string]$MetadataPath = $null

    if (!$WebConfigPath)
    {
        $WebConfigPath = Get-AX7DeploymentAosWebConfigPath
    }
    $MetadataPath = Get-AX7DeploymentAosWebConfigSetting -WebConfigPath $WebConfigPath -Name "Aos.MetadataDirectory"

    return $MetadataPath
}

<#
.SYNOPSIS
    Get the Dynamics AX 7.0 deployment packages path.

.DESCRIPTION
    Value is extracted from the specified web.config file path of the default
    AOS web config if no web config path is specified.
#>
function Get-AX7DeploymentPackagesPath
{
    [Cmdletbinding()]
    Param([string]$WebConfigPath)
    
    # Get verbose preference from caller.
    $VerbosePreference = $PSCmdlet.GetVariableValue("VerbosePreference")
        
    [string]$PackagesPath = $null

    if (!$WebConfigPath)
    {
        $WebConfigPath = Get-AX7DeploymentAosWebConfigPath
    }
    $PackagesPath = Get-AX7DeploymentAosWebConfigSetting -WebConfigPath $WebConfigPath -Name "Aos.PackageDirectory"

    return $PackagesPath
}

<#
.SYNOPSIS
    Get the Dynamics AX 7.0 deployment database name.

.DESCRIPTION
    Value is extracted from the specified web.config file path of the default
    AOS web config if no web config path is specified.
#>
function Get-AX7DeploymentDatabaseName
{
    [Cmdletbinding()]
    Param([string]$WebConfigPath)
    
    # Get verbose preference from caller.
    $VerbosePreference = $PSCmdlet.GetVariableValue("VerbosePreference")
        
    [string]$DatabaseName = $null

    if (!$WebConfigPath)
    {
        $WebConfigPath = Get-AX7DeploymentAosWebConfigPath
    }
    $DatabaseName = Get-AX7DeploymentAosWebConfigSetting -WebConfigPath $WebConfigPath -Name "DataAccess.Database"

    return $DatabaseName
}

<#
.SYNOPSIS
    Get the Dynamics AX 7.0 deployment database server.

.DESCRIPTION
    Value is extracted from the specified web.config file path of the default
    AOS web config if no web config path is specified.
#>
function Get-AX7DeploymentDatabaseServer
{
    [Cmdletbinding()]
    Param([string]$WebConfigPath)
    
    # Get verbose preference from caller.
    $VerbosePreference = $PSCmdlet.GetVariableValue("VerbosePreference")
        
    [string]$DatabaseServer = $null

    if (!$WebConfigPath)
    {
        $WebConfigPath = Get-AX7DeploymentAosWebConfigPath
    }
    $DatabaseServer = Get-AX7DeploymentAosWebConfigSetting -WebConfigPath $WebConfigPath -Name "DataAccess.DbServer"

    return $DatabaseServer
}

<#
.SYNOPSIS
    Stop the Dynamics AX 7.0 deployment AOS website.
#>
function Stop-AX7DeploymentAosWebsite
{
    [Cmdletbinding()]
    Param([string]$WebsiteName)
    
    # Get verbose preference from caller.
    $VerbosePreference = $PSCmdlet.GetVariableValue("VerbosePreference")

    # Get website and stop it if not already stopped.
    $Website = Get-AX7DeploymentAosWebsite -WebsiteName $WebsiteName
    if ($Website)
    {
        Write-Message "- AOS website state: $($Website.State)" -Diag

        # State is empty if IIS is not running in which case the website is already stopped.
        if ($Website.State -and $Website.State -ine "Stopped")
        {
            Write-Message "- Stopping AOS website..." -Diag
            $Website.Stop()
            Write-Message "- AOS website state after stop: $($Website.State)" -Diag
        }
    }

    # If IIS Express instances are running, stop those
    $expressSites = Get-Process | Where-Object { $_.Name -eq "iisexpress" }
    if ($expressSites.Length -gt 0)
    {
        Write-Message "- Stopping IIS Express instances"
        foreach($site in $expressSites)
        {
            Stop-Process $site -Force
        }
        Write-Message "- IIS Express instances stopped"
    }
}

<#
.SYNOPSIS
    Start the Dynamics AX 7.0 deployment AOS website.
#>
function Start-AX7DeploymentAosWebsite
{
    [Cmdletbinding()]
    Param([string]$WebsiteName)

    # Get verbose preference from caller.
    $VerbosePreference = $PSCmdlet.GetVariableValue("VerbosePreference")

    # Get website and start it if not already started.
    $Website = Get-AX7DeploymentAosWebsite -WebsiteName $WebsiteName
    if ($Website)
    {
        Write-Message "- AOS website state: $($Website.State)" -Diag
        # State is empty if IIS is not running.
        if (!($Website.State))
        {
            Start-IIS
            Write-Message "- AOS website state after IIS start: $($Website.State)" -Diag
        }

        if ($Website.State -and $Website.State -ine "Started")
        {
            Write-Message "- Starting AOS website..." -Diag
            $Website.Start()
            Write-Message "- AOS website state after start: $($Website.State)" -Diag
        }
    }
}

<#
.SYNOPSIS
    Stop the Dynamics AX 7.0 services and IIS.
#>
function Stop-AX7Deployment
{
    [Cmdletbinding()]
    Param([int]$ServiceStopWaitSec = 30, [int]$ProcessStopWaitSec = 30)

    # Get verbose preference from caller.
    $VerbosePreference = $PSCmdlet.GetVariableValue("VerbosePreference")

    # There are a number of Dynamics web sites. Safer to stop IIS completely.
    Stop-IIS

    $DynamicsServiceNames = @("DynamicsAxBatch", "Microsoft.Dynamics.AX.Framework.Tools.DMF.SSISHelperService.exe", "MR2012ProcessService")
    foreach ($DynamicsServiceName in $DynamicsServiceNames)
    {
        $Service = Get-Service -Name $DynamicsServiceName -ErrorAction SilentlyContinue
        if ($Service)
        {
            if ($Service.Status -ne [System.ServiceProcess.ServiceControllerStatus]::Stopped)
            {
                # Get the service process ID to track if it has exited when the service has stopped.
                [UInt32]$ServiceProcessId = 0
                $WmiService = Get-WmiObject -Class Win32_Service -Filter "Name = '$($Service.ServiceName)'" -ErrorAction SilentlyContinue
                if ($WmiService)
                {
                    if ($WmiService.ProcessId -gt $ServiceProcessId)
                    {
                        $ServiceProcessId = $WmiService.ProcessId
                        Write-Message "- The $($DynamicsServiceName) service has process ID: $ServiceProcessId" -Diag
                    }
                    else
                    {
                        Write-Message "- The $($DynamicsServiceName) service does not have a process ID." -Diag
                    }
                }
                else
                {
                    Write-Message "- No $($Service.ServiceName) service found through WMI. Cannot get process ID of the service." -Warning
                }

                # Signal the service to stop.
                Write-Message "- Stopping the $($DynamicsServiceName) service (Status: $($Service.Status))..." -Diag
                Stop-Service -Name $DynamicsServiceName

                # Wait for the service to stop.
                if ($ServiceStopWaitSec -gt 0)
                {
                    Write-Message "- Waiting up to $($ServiceStopWaitSec) seconds for the $($DynamicsServiceName) service to stop (Status: $($Service.Status))..." -Diag
                    # This will throw a System.ServiceProcess.TimeoutException if the stopped state is not reached within the timeout.
                    $Service.WaitForStatus([System.ServiceProcess.ServiceControllerStatus]::Stopped, [TimeSpan]::FromSeconds($ServiceStopWaitSec))
                    Write-Message "- The $($DynamicsServiceName) service has been stopped (Status: $($Service.Status))." -Diag
                }

                # Wait for the process, if any was found, to exit.
                if ($ProcessStopWaitSec -gt 0 -and $ServiceProcessId -gt 0)
                {
                    # If the process is found, wait for it to exit.
                    $ServiceProcess = Get-Process -Id $ServiceProcessId -ErrorAction SilentlyContinue
                    if ($ServiceProcess)
                    {
                        Write-Message "- Waiting up to $($ProcessStopWaitSec) seconds for the $($DynamicsServiceName) service process ID $($ServiceProcessId) to exit..." -Diag
                        # This will throw a System.TimeoutException if the process does not exit within the timeout.
                        Wait-Process -Id $ServiceProcessId -Timeout $ProcessStopWaitSec
                    }
                    Write-Message "- The $($DynamicsServiceName) service process ID $($ServiceProcessId) has exited." -Diag
                }
            }
            else
            {
                Write-Message "- The $($DynamicsServiceName) service is already stopped." -Diag
            }
        }
        else
        {
            Write-Message "- No $($DynamicsServiceName) service found." -Diag
        }
    }
}

<#
.SYNOPSIS
    Start the Dynamics AX 7.0 deployment services and IIS.
#>
function Start-AX7Deployment
{
    [Cmdletbinding()]
    Param()

    # Get verbose preference from caller.
    $VerbosePreference = $PSCmdlet.GetVariableValue("VerbosePreference")

    $DynamicsServiceNames = @("DynamicsAxBatch", "Microsoft.Dynamics.AX.Framework.Tools.DMF.SSISHelperService.exe", "MR2012ProcessService")
    foreach ($DynamicsServiceName in $DynamicsServiceNames)
    {
        $Service = Get-Service -Name $DynamicsServiceName -ErrorAction SilentlyContinue
        if ($Service)
        {
            if ($Service.Status -ne [System.ServiceProcess.ServiceControllerStatus]::Running)
            {
                Write-Message "- Starting $($DynamicsServiceName) service..." -Diag
                Start-Service -Name $DynamicsServiceName
                Write-Message "- $($DynamicsServiceName) service successfully started." -Diag
            }
            else
            {
                Write-Message "- $($DynamicsServiceName) service is already running." -Diag
            }
        }
        else
        {
            Write-Message "- No $($DynamicsServiceName) service found." -Diag
        }
    }

    # Start IIS back up.
    Start-IIS
}

# Functions to export from this module (sorted alphabetically).
$ExportFunctions = @(
    "Get-AX7DeploymentAosWebConfigPath",
    "Get-AX7DeploymentAosWifConfigPath",
    "Get-AX7SdkBackupPath",
    "Get-AX7SdkDeploymentAosWebsiteName",
    "Get-AX7SdkDeploymentBinariesPath",
    "Get-AX7SdkDeploymentDatabaseName",
    "Get-AX7SdkDeploymentDatabaseServer",
    "Get-AX7SdkDeploymentMetadataPath",
    "Get-AX7SdkDeploymentPackagesPath",
    "Get-AX7SdkPath",
    "Get-AX7SdkTeamFoundationServerUrl",
    "Set-AX7SdkEnvironmentVariables",
    "Set-AX7SdkRegistryValues",
    "Set-AX7SdkRegistryValuesFromAosWebConfig",
    "Start-AX7Deployment",
    "Start-AX7DeploymentAosWebsite",
    "Stop-AX7Deployment",
    "Stop-AX7DeploymentAosWebsite",
    "Write-Message"
)

Export-ModuleMember -Function $ExportFunctions
# SIG # Begin signature block
# MIIkaQYJKoZIhvcNAQcCoIIkWjCCJFYCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBfn7ZIHE/dU5JQ
# FRh8xGVvtqdkDd/YaMWaWJ3CN9H7pKCCDYEwggX/MIID56ADAgECAhMzAAABA14l
# HJkfox64AAAAAAEDMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMTgwNzEyMjAwODQ4WhcNMTkwNzI2MjAwODQ4WjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQDRlHY25oarNv5p+UZ8i4hQy5Bwf7BVqSQdfjnnBZ8PrHuXss5zCvvUmyRcFrU5
# 3Rt+M2wR/Dsm85iqXVNrqsPsE7jS789Xf8xly69NLjKxVitONAeJ/mkhvT5E+94S
# nYW/fHaGfXKxdpth5opkTEbOttU6jHeTd2chnLZaBl5HhvU80QnKDT3NsumhUHjR
# hIjiATwi/K+WCMxdmcDt66VamJL1yEBOanOv3uN0etNfRpe84mcod5mswQ4xFo8A
# DwH+S15UD8rEZT8K46NG2/YsAzoZvmgFFpzmfzS/p4eNZTkmyWPU78XdvSX+/Sj0
# NIZ5rCrVXzCRO+QUauuxygQjAgMBAAGjggF+MIIBejAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQUR77Ay+GmP/1l1jjyA123r3f3QP8w
# UAYDVR0RBEkwR6RFMEMxKTAnBgNVBAsTIE1pY3Jvc29mdCBPcGVyYXRpb25zIFB1
# ZXJ0byBSaWNvMRYwFAYDVQQFEw0yMzAwMTIrNDM3OTY1MB8GA1UdIwQYMBaAFEhu
# ZOVQBdOCqhc3NyK1bajKdQKVMFQGA1UdHwRNMEswSaBHoEWGQ2h0dHA6Ly93d3cu
# bWljcm9zb2Z0LmNvbS9wa2lvcHMvY3JsL01pY0NvZFNpZ1BDQTIwMTFfMjAxMS0w
# Ny0wOC5jcmwwYQYIKwYBBQUHAQEEVTBTMFEGCCsGAQUFBzAChkVodHRwOi8vd3d3
# Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NlcnRzL01pY0NvZFNpZ1BDQTIwMTFfMjAx
# MS0wNy0wOC5jcnQwDAYDVR0TAQH/BAIwADANBgkqhkiG9w0BAQsFAAOCAgEAn/XJ
# Uw0/DSbsokTYDdGfY5YGSz8eXMUzo6TDbK8fwAG662XsnjMQD6esW9S9kGEX5zHn
# wya0rPUn00iThoj+EjWRZCLRay07qCwVlCnSN5bmNf8MzsgGFhaeJLHiOfluDnjY
# DBu2KWAndjQkm925l3XLATutghIWIoCJFYS7mFAgsBcmhkmvzn1FFUM0ls+BXBgs
# 1JPyZ6vic8g9o838Mh5gHOmwGzD7LLsHLpaEk0UoVFzNlv2g24HYtjDKQ7HzSMCy
# RhxdXnYqWJ/U7vL0+khMtWGLsIxB6aq4nZD0/2pCD7k+6Q7slPyNgLt44yOneFuy
# bR/5WcF9ttE5yXnggxxgCto9sNHtNr9FB+kbNm7lPTsFA6fUpyUSj+Z2oxOzRVpD
# MYLa2ISuubAfdfX2HX1RETcn6LU1hHH3V6qu+olxyZjSnlpkdr6Mw30VapHxFPTy
# 2TUxuNty+rR1yIibar+YRcdmstf/zpKQdeTr5obSyBvbJ8BblW9Jb1hdaSreU0v4
# 6Mp79mwV+QMZDxGFqk+av6pX3WDG9XEg9FGomsrp0es0Rz11+iLsVT9qGTlrEOla
# P470I3gwsvKmOMs1jaqYWSRAuDpnpAdfoP7YO0kT+wzh7Qttg1DO8H8+4NkI6Iwh
# SkHC3uuOW+4Dwx1ubuZUNWZncnwa6lL2IsRyP64wggd6MIIFYqADAgECAgphDpDS
# AAAAAAADMA0GCSqGSIb3DQEBCwUAMIGIMQswCQYDVQQGEwJVUzETMBEGA1UECBMK
# V2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
# IENvcnBvcmF0aW9uMTIwMAYDVQQDEylNaWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0
# ZSBBdXRob3JpdHkgMjAxMTAeFw0xMTA3MDgyMDU5MDlaFw0yNjA3MDgyMTA5MDla
# MH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdS
# ZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMT
# H01pY3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBIDIwMTEwggIiMA0GCSqGSIb3DQEB
# AQUAA4ICDwAwggIKAoICAQCr8PpyEBwurdhuqoIQTTS68rZYIZ9CGypr6VpQqrgG
# OBoESbp/wwwe3TdrxhLYC/A4wpkGsMg51QEUMULTiQ15ZId+lGAkbK+eSZzpaF7S
# 35tTsgosw6/ZqSuuegmv15ZZymAaBelmdugyUiYSL+erCFDPs0S3XdjELgN1q2jz
# y23zOlyhFvRGuuA4ZKxuZDV4pqBjDy3TQJP4494HDdVceaVJKecNvqATd76UPe/7
# 4ytaEB9NViiienLgEjq3SV7Y7e1DkYPZe7J7hhvZPrGMXeiJT4Qa8qEvWeSQOy2u
# M1jFtz7+MtOzAz2xsq+SOH7SnYAs9U5WkSE1JcM5bmR/U7qcD60ZI4TL9LoDho33
# X/DQUr+MlIe8wCF0JV8YKLbMJyg4JZg5SjbPfLGSrhwjp6lm7GEfauEoSZ1fiOIl
# XdMhSz5SxLVXPyQD8NF6Wy/VI+NwXQ9RRnez+ADhvKwCgl/bwBWzvRvUVUvnOaEP
# 6SNJvBi4RHxF5MHDcnrgcuck379GmcXvwhxX24ON7E1JMKerjt/sW5+v/N2wZuLB
# l4F77dbtS+dJKacTKKanfWeA5opieF+yL4TXV5xcv3coKPHtbcMojyyPQDdPweGF
# RInECUzF1KVDL3SV9274eCBYLBNdYJWaPk8zhNqwiBfenk70lrC8RqBsmNLg1oiM
# CwIDAQABo4IB7TCCAekwEAYJKwYBBAGCNxUBBAMCAQAwHQYDVR0OBBYEFEhuZOVQ
# BdOCqhc3NyK1bajKdQKVMBkGCSsGAQQBgjcUAgQMHgoAUwB1AGIAQwBBMAsGA1Ud
# DwQEAwIBhjAPBgNVHRMBAf8EBTADAQH/MB8GA1UdIwQYMBaAFHItOgIxkEO5FAVO
# 4eqnxzHRI4k0MFoGA1UdHwRTMFEwT6BNoEuGSWh0dHA6Ly9jcmwubWljcm9zb2Z0
# LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01pY1Jvb0NlckF1dDIwMTFfMjAxMV8wM18y
# Mi5jcmwwXgYIKwYBBQUHAQEEUjBQME4GCCsGAQUFBzAChkJodHRwOi8vd3d3Lm1p
# Y3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY1Jvb0NlckF1dDIwMTFfMjAxMV8wM18y
# Mi5jcnQwgZ8GA1UdIASBlzCBlDCBkQYJKwYBBAGCNy4DMIGDMD8GCCsGAQUFBwIB
# FjNodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2RvY3MvcHJpbWFyeWNw
# cy5odG0wQAYIKwYBBQUHAgIwNB4yIB0ATABlAGcAYQBsAF8AcABvAGwAaQBjAHkA
# XwBzAHQAYQB0AGUAbQBlAG4AdAAuIB0wDQYJKoZIhvcNAQELBQADggIBAGfyhqWY
# 4FR5Gi7T2HRnIpsLlhHhY5KZQpZ90nkMkMFlXy4sPvjDctFtg/6+P+gKyju/R6mj
# 82nbY78iNaWXXWWEkH2LRlBV2AySfNIaSxzzPEKLUtCw/WvjPgcuKZvmPRul1LUd
# d5Q54ulkyUQ9eHoj8xN9ppB0g430yyYCRirCihC7pKkFDJvtaPpoLpWgKj8qa1hJ
# Yx8JaW5amJbkg/TAj/NGK978O9C9Ne9uJa7lryft0N3zDq+ZKJeYTQ49C/IIidYf
# wzIY4vDFLc5bnrRJOQrGCsLGra7lstnbFYhRRVg4MnEnGn+x9Cf43iw6IGmYslmJ
# aG5vp7d0w0AFBqYBKig+gj8TTWYLwLNN9eGPfxxvFX1Fp3blQCplo8NdUmKGwx1j
# NpeG39rz+PIWoZon4c2ll9DuXWNB41sHnIc+BncG0QaxdR8UvmFhtfDcxhsEvt9B
# xw4o7t5lL+yX9qFcltgA1qFGvVnzl6UJS0gQmYAf0AApxbGbpT9Fdx41xtKiop96
# eiL6SJUfq/tHI4D1nvi/a7dLl+LrdXga7Oo3mXkYS//WsyNodeav+vyL6wuA6mk7
# r/ww7QRMjt/fdW1jkT3RnVZOT7+AVyKheBEyIXrvQQqxP/uozKRdwaGIm1dxVk5I
# RcBCyZt2WwqASGv9eZ/BvW1taslScxMNelDNMYIWPjCCFjoCAQEwgZUwfjELMAkG
# A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
# HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYGA1UEAxMfTWljcm9z
# b2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMQITMwAAAQNeJRyZH6MeuAAAAAABAzAN
# BglghkgBZQMEAgEFAKCB0DAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgor
# BgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgHzV5mQhH
# NinDsBVsH6mXVZA41hEHBBdPRuSMbB0edi8wZAYKKwYBBAGCNwIBDDFWMFSgNoA0
# AGcAbABvAGIAYQBsAGkAegBlAC4AYwB1AGwAdAB1AHIAZQAuAG4AYgAtAE4ATwAu
# AGoAc6EagBhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20wDQYJKoZIhvcNAQEBBQAE
# ggEAecnObQn/5MRjfE6TpTbnMKXcAvsdcrpTzVzp81XSWnxR+7VUTZ2MC9/XZrqR
# s7m3gY9ufTtK5zo9HjDdJb9ohsQlCs8KN0uNM1Dq+MeMcCSsydKBkqnhzt95b9E7
# g3G6BRaU7fH1SCESQaytWTUTo1c3CA9Mx1UAci0tqtS/FMOB4wz/R/T3ePSL3iDk
# 4r70lsnyFQBTF/sZx+evShCJVp23eeQhsGt9TiJsU2UyD+6FhFrho1QMpvHPMG4r
# jPdAUykuCh+50inE5heVRMdAHZQpZtO/BNqFULzDIotfhqaFELN5jusJ7KQbuwpf
# mlv6J2yC8v6HMaRsvUHBQNRrS6GCE6YwghOiBgorBgEEAYI3AwMBMYITkjCCE44G
# CSqGSIb3DQEHAqCCE38wghN7AgEDMQ8wDQYJYIZIAWUDBAIBBQAwggFUBgsqhkiG
# 9w0BCRABBKCCAUMEggE/MIIBOwIBAQYKKwYBBAGEWQoDATAxMA0GCWCGSAFlAwQC
# AQUABCCqjQqhgNqHuQnWBw5RJUJWAd1vEyOidXEOGO0nL8x48gIGW60mvS4iGBMy
# MDE4MDkzMDE1NTA0OC4wMjRaMAcCAQGAAgH0oIHQpIHNMIHKMQswCQYDVQQGEwJV
# UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UE
# ChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQLExxNaWNyb3NvZnQgQW1l
# cmljYSBPcGVyYXRpb25zMSYwJAYDVQQLEx1UaGFsZXMgVFNTIEVTTjpDM0IwLTBG
# NkEtNDExMTElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2VydmljZaCC
# DxIwggZxMIIEWaADAgECAgphCYEqAAAAAAACMA0GCSqGSIb3DQEBCwUAMIGIMQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMTIwMAYDVQQDEylNaWNy
# b3NvZnQgUm9vdCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHkgMjAxMDAeFw0xMDA3MDEy
# MTM2NTVaFw0yNTA3MDEyMTQ2NTVaMHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpX
# YXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQg
# Q29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAy
# MDEwMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAqR0NvHcRijog7PwT
# l/X6f2mUa3RUENWlCgCChfvtfGhLLF/Fw+Vhwna3PmYrW/AVUycEMR9BGxqVHc4J
# E458YTBZsTBED/FgiIRUQwzXTbg4CLNC3ZOs1nMwVyaCo0UN0Or1R4HNvyRgMlhg
# RvJYR4YyhB50YWeRX4FUsc+TTJLBxKZd0WETbijGGvmGgLvfYfxGwScdJGcSchoh
# iq9LZIlQYrFd/XcfPfBXday9ikJNQFHRD5wGPmd/9WbAA5ZEfu/QS/1u5ZrKsajy
# eioKMfDaTgaRtogINeh4HLDpmc085y9Euqf03GS9pAHBIAmTeM38vMDJRF1eFpwB
# BU8iTQIDAQABo4IB5jCCAeIwEAYJKwYBBAGCNxUBBAMCAQAwHQYDVR0OBBYEFNVj
# OlyKMZDzQ3t8RhvFM2hahW1VMBkGCSsGAQQBgjcUAgQMHgoAUwB1AGIAQwBBMAsG
# A1UdDwQEAwIBhjAPBgNVHRMBAf8EBTADAQH/MB8GA1UdIwQYMBaAFNX2VsuP6KJc
# YmjRPZSQW9fOmhjEMFYGA1UdHwRPME0wS6BJoEeGRWh0dHA6Ly9jcmwubWljcm9z
# b2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01pY1Jvb0NlckF1dF8yMDEwLTA2LTIz
# LmNybDBaBggrBgEFBQcBAQROMEwwSgYIKwYBBQUHMAKGPmh0dHA6Ly93d3cubWlj
# cm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljUm9vQ2VyQXV0XzIwMTAtMDYtMjMuY3J0
# MIGgBgNVHSABAf8EgZUwgZIwgY8GCSsGAQQBgjcuAzCBgTA9BggrBgEFBQcCARYx
# aHR0cDovL3d3dy5taWNyb3NvZnQuY29tL1BLSS9kb2NzL0NQUy9kZWZhdWx0Lmh0
# bTBABggrBgEFBQcCAjA0HjIgHQBMAGUAZwBhAGwAXwBQAG8AbABpAGMAeQBfAFMA
# dABhAHQAZQBtAGUAbgB0AC4gHTANBgkqhkiG9w0BAQsFAAOCAgEAB+aIUQ3ixuCY
# P4FxAz2do6Ehb7Prpsz1Mb7PBeKp/vpXbRkws8LFZslq3/Xn8Hi9x6ieJeP5vO1r
# VFcIK1GCRBL7uVOMzPRgEop2zEBAQZvcXBf/XPleFzWYJFZLdO9CEMivv3/Gf/I3
# fVo/HPKZeUqRUgCvOA8X9S95gWXZqbVr5MfO9sp6AG9LMEQkIjzP7QOllo9ZKby2
# /QThcJ8ySif9Va8v/rbljjO7Yl+a21dA6fHOmWaQjP9qYn/dxUoLkSbiOewZSnFj
# nXshbcOco6I8+n99lmqQeKZt0uGc+R38ONiU9MalCpaGpL2eGq4EQoO4tYCbIjgg
# tSXlZOz39L9+Y1klD3ouOVd2onGqBooPiRa6YacRy5rYDkeagMXQzafQ732D8OE7
# cQnfXXSYIghh2rBQHm+98eEA3+cxB6STOvdlR3jo+KhIq/fecn5ha293qYHLpwms
# ObvsxsvYgrRyzR30uIUBHoD7G4kqVDmyW9rIDVWZeodzOwjmmC3qjeAzLhIp9cAv
# VCch98isTtoouLGp25ayp0Kiyc8ZQU3ghvkqmqMRZjDTu3QyS99je/WZii8bxyGv
# WbWu3EQ8l1Bx16HSxVXjad5XwdHeMMD9zOZN+w2/XU/pnR4ZOC+8z1gFLu8NoFA1
# 2u8JJxzVs341Hgi62jbb01+P3nSISRIwggTxMIID2aADAgECAhMzAAAAyCQZy6pU
# 5mwqAAAAAADIMA0GCSqGSIb3DQEBCwUAMHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBD
# QSAyMDEwMB4XDTE4MDgyMzIwMjYxM1oXDTE5MTEyMzIwMjYxM1owgcoxCzAJBgNV
# BAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4w
# HAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJTAjBgNVBAsTHE1pY3Jvc29m
# dCBBbWVyaWNhIE9wZXJhdGlvbnMxJjAkBgNVBAsTHVRoYWxlcyBUU1MgRVNOOkMz
# QjAtMEY2QS00MTExMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2
# aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA44mGILYaJ4ThEuFw
# A78ZmH2oQSIy8wqnqQJ5ic7nvGK7gQKjxv4bPrRifeFjzps3E2TG62FTBWst38CA
# l/s9ZozfsFt3/Crv5hHOHVbHkVYrBH3jr/xi9JXLVm9Ub6LDTEB7l9F4mj4q8HIV
# As2YP1h7x1TZwlbdXmdNHzF5EfmF+y6KeeAvskT89W0y1qTyJ0RRYsGD3uKqAlF+
# aYbZ7HOUsGFMr7H/MtYcYMLDyTLq5QPpjEFJ+27yqGyvyfj7m/dhbV7IaYdAW6wZ
# QNjjMWnPk77xxFjLpgihGuWMK14vvyCBAzzumLFPO5+KC2RpfYMNV+iNRKV998js
# a1ZwpwIDAQABo4IBGzCCARcwHQYDVR0OBBYEFP4w0enyVikOI9BFM3UMpZ7CUkDA
# MB8GA1UdIwQYMBaAFNVjOlyKMZDzQ3t8RhvFM2hahW1VMFYGA1UdHwRPME0wS6BJ
# oEeGRWh0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01p
# Y1RpbVN0YVBDQV8yMDEwLTA3LTAxLmNybDBaBggrBgEFBQcBAQROMEwwSgYIKwYB
# BQUHMAKGPmh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljVGlt
# U3RhUENBXzIwMTAtMDctMDEuY3J0MAwGA1UdEwEB/wQCMAAwEwYDVR0lBAwwCgYI
# KwYBBQUHAwgwDQYJKoZIhvcNAQELBQADggEBAJOaM61oDQ2og12OxdrFda7V5H2N
# kJRuxSPqVpMHgz11e4RGX+iHYXRExePOni94SaZzSGF3mbjfrtlUULHqvgyYWVT1
# 7dRSbtJqoGZnJcUPbo5MhosHij/ogKobX8YKOSprcSwjdIB5MFPUXZyO/IC8RaRt
# 0ksc+GjOnX+FYU8RL789SmlLMjiXc2eNqzJ2YnaNkKg0O5syPla2HxLceAZuNQYB
# 5zUMsvRTtWQuxrGRd5IGuraPDnwHtU1rxLj+aCwmmrPLJOWXulzMxLgZMrNxZdNr
# Dg+YRBdeepWhPMT68IY4ql5G/oqkN2lFxCt/tZEfDAMkBJ4ICySs+A55NzOhggOk
# MIICjAIBATCB+qGB0KSBzTCByjELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFtZXJpY2EgT3BlcmF0aW9uczEm
# MCQGA1UECxMdVGhhbGVzIFRTUyBFU046QzNCMC0wRjZBLTQxMTExJTAjBgNVBAMT
# HE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2WiJQoBATAJBgUrDgMCGgUAAxUA
# NJcmcaPsEmrOgrfryoIY7WPSXU2ggdowgdekgdQwgdExCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xJTAjBgNVBAsTHE1pY3Jvc29mdCBBbWVyaWNh
# IE9wZXJhdGlvbnMxJzAlBgNVBAsTHm5DaXBoZXIgTlRTIEVTTjoyNjY1LTRDM0Yt
# QzVERTErMCkGA1UEAxMiTWljcm9zb2Z0IFRpbWUgU291cmNlIE1hc3RlciBDbG9j
# azANBgkqhkiG9w0BAQUFAAIFAN9bGTcwIhgPMjAxODA5MzAwOTQzMTlaGA8yMDE4
# MTAwMTA5NDMxOVowczA5BgorBgEEAYRZCgQBMSswKTAKAgUA31sZNwIBADAGAgEA
# AgEfMAcCAQACAhUeMAoCBQDfXGq3AgEAMDYGCisGAQQBhFkKBAIxKDAmMAwGCisG
# AQQBhFkKAwGgCjAIAgEAAgMW42ChCjAIAgEAAgMehIAwDQYJKoZIhvcNAQEFBQAD
# ggEBAFcuBiTipn86w9hzpMLwSnxCvK5VtlL6xuCqSPIOZJQJ6NjxCQwEC2iya3DB
# k6a7ZDBG/xQBQ4SXL6z5/qbfhksfL6iLPHXaZ5ZeeAunfa7szHz9nro8yIUC229U
# l8I01/Q34WnBRZisQQJ0CeZ4+Cy7ZqY49tIt1Z9k5NyKVuBvxgv7zOM4Xda96z3y
# g/Npog82KPo/Qt9oS6/uPMHJFvkPCGXfLWL/1BcfV7t+pUVEeQFP+2La4uU1QEvf
# k/l/UcvbDFG74WMLFXkvCZx+B7uX7IyIKfxen8y7I/xs2SsorE9CLMIhcLygkLC6
# ibdLp0hqK+gK2Zz5gas2MzoBMQQxggL1MIIC8QIBATCBkzB8MQswCQYDVQQGEwJV
# UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UE
# ChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGlt
# ZS1TdGFtcCBQQ0EgMjAxMAITMwAAAMgkGcuqVOZsKgAAAAAAyDANBglghkgBZQME
# AgEFAKCCATIwGgYJKoZIhvcNAQkDMQ0GCyqGSIb3DQEJEAEEMC8GCSqGSIb3DQEJ
# BDEiBCCn+v/APwU973q7xlrYLcrsRxsarYkHxqwnzv1I0rcvKzCB4gYLKoZIhvcN
# AQkQAgwxgdIwgc8wgcwwgbEEFDSXJnGj7BJqzoK368qCGO1j0l1NMIGYMIGApH4w
# fDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1Jl
# ZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMd
# TWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTACEzMAAADIJBnLqlTmbCoAAAAA
# AMgwFgQU28Y57J0gT1zSaMbYIWDK3UuYm70wDQYJKoZIhvcNAQELBQAEggEAgaAO
# NBclupbZ1chfoywUuo6UCE3nvdh/k/DJiVU7KcWey8rPcxdx/wHwDwdRxjqNgvLd
# nH384tpveKznoyV9joXR7dy8JAaB7ZSoiholuNEHrPBYA1yKS1Nrce//eqRwOd77
# iZcQfML13ZF3+xstQwQoCPSAX/HDHhbR87Q+TMAcgZmj3chznMVqIl/h91uGeuUM
# lZtndPa0K2NKfYPHzuX4KMz34Scs5gB79fCI6xxf0Ybwie+dwOg2TB0ieOeOe4Dh
# +AIt8+e0jXJhk0GgqpLhRdOc5X93RgIPHfJbDZh7GfKHSEkx1Al9x8gwrsUhljwU
# ifQaAYjz/eTY5xs0yQ==
# SIG # End signature block

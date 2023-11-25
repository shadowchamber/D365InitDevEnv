[Cmdletbinding()]
Param(
	[Parameter(HelpMessage="Current machine's system drive")]
	[string]$SystemDrive = $env:SystemDrive,
	[Parameter(HelpMessage="Current machine's service drive")]
	[string]$ServiceDrive = "K:",
	[Parameter(HelpMessage="The folder to host git repositories")]
	[string]$RepoDir = "$SystemDrive\Repos",
	[Parameter(HelpMessage="The time zone to use")]
	[string]$TimeZone = "Central European Standard Time",
	[Parameter(HelpMessage="The locale to use")]
	[string]$Locale = "en-us",
	[Parameter(HelpMessage="Determine whether to install additional software such as Chrome and Notepad++")]
	[switch]$InstallSoftware
)

Import-Module $PSScriptRoot\Init-Env.psm1

# Set timezone
Write-Host "Setting time zone $TimeZone"
Set-TimeZone -Name $TimeZone

# Set user locale
Write-Host "Setting current locale to $Locale"
Set-WinUserLanguageList -LanguageList $Locale -Force

# Install chocolatey
if (($InstallSoftware -eq $true) -or ($CloneRepos -eq $true))
{
	Set-ExecutionPolicy Bypass -Scope Process -Force; iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
}

if ($InstallSoftware -eq $true)
{
	# Install Google Chrome and Edge
	Install-ChocoPackage -PackageName googlechrome
	Install-ChocoPackage -PackageName microsoft-edge

	# Install Notepad++
	Install-ChocoPackage -PackageName notepadplusplus
	
	# Install DocFx relevant software
	Install-ChocoPackage -PackageName docfx
	Install-ChocoPackage -PackageName wkhtmltopdf

	# Additional useful software
	Install-ChocoPackage -PackageName 7zip
	Install-ChocoPackage -PackageName postman
	Install-ChocoPackage -PackageName microsoftazurestorageexplorer
	Install-ChocoPackage -PackageName vscode
	Install-ChocoPackage -PackageName nodejs
	Install-ChocoPackage -PackageName typescript
	Install-ChocoPackage -PackageName azure-cli
	Install-ChocoPackage -PackageName dotnetcore-sdk -Version 2.1.503
	Install-ChocoPackage -PackageName ServiceBusExplorer

	# Create desktop shortcut to ServiceBusExplorer
    Create-Shortcut -ShortcutPath "$Home\Desktop\Service Bus Explorer.lnk" "$env:ChocolateyInstall\lib\ServiceBusExplorer\tools\ServiceBusExplorer.exe"	

    # Install git for Windows
	Install-ChocoPackage -PackageName git.install

	# Install git Extensions
	Install-ChocoPackage -PackageName gitextensions

	# Install git Fork
	Install-ChocoPackage -PackageName git-fork

    # Install tortoisegit
    Install-ChocoPackage -PackageName tortoisegit

    # Install github desktop
    Install-ChocoPackage -PackageName github-desktop
}
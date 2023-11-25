function Install-ChocoPackage
{
    Param(
        [Parameter(Mandatory=$true)]
        [string]$PackageName,
        [string]$Version = ""
    )

    $ChocoLibPath = "C:\ProgramData\chocolatey\lib"

    $start_time = Get-Date

    if(-not(test-path $ChocoLibPath)) {
        Write-Host "[INFO]Installing $PackageName..." -ForegroundColor Yellow       

        if ($Version -eq "")
        {
            choco install $PackageName --yes
        }
        else
        {
            choco install $PackageName --yes --version $Version
        }
    }
    else {
        Write-Host "[INFO]Upgrading $PackageName..." -ForegroundColor Yellow

        if ($Version -eq "")
        {
            choco upgrade $PackageName --yes
        }
        else
        {
            choco upgrade $PackageName --yes --version $Version
        }
    }

    Write-Host "Time taken: $((Get-Date).Subtract($start_time).Seconds) second(s)"
}

function Add-Shortcut
{
    Param(
        [Parameter(Mandatory=$true)]
        [string]$ShortcutPath,
        [Parameter(Mandatory=$true)]
        [string]$TargetPath
    )    
    if ((Test-Path -Path $TargetPath) -eq $true)
    {
        $WshShell = New-Object -comObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut("$ShortcutPath")
        $Shortcut.TargetPath = "$TargetPath"
        $Shortcut.Save()
    }
}

function Add-QuickAccessFolders
{
    Param(
        [Parameter(Mandatory=$true)]
        [string]$FolderList
    )
    Write-Host "Pinning folders to quick access:"
    $FolderArray = $FolderList.Split(",")
    foreach ($Folder in $FolderArray)
    {
        if ((Test-Path $Folder) -eq $true)
        {
            Write-Host "Folder $Folder"
            $Object = New-Object -ComObject Shell.Application
            $Object.Namespace($Folder).Self.InvokeVerb("pintohome")
        }
    }
}

function Copy-PSModules
{
	Param(
		[string]$ModulesPath,
		[string]$DestinationPath = "C:\Program Files\WindowsPowerShell\Modules"
	)
	
	$Source = Join-Path -Path $ModulesPath -ChildPath "*"
	Copy-Item -Path $Source -Destination $DestinationPath -Recurse -Force
}

function Copy-PSProfile
{
	Param(
		[string]$SourcePath,
		[string]$DestinationPath = "C:\Windows\System32\WindowsPowerShell\v1.0"
	)
	
	Copy-Item -Path $SourcePath -Destination $DestinationPath -Force
}

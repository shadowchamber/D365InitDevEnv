function Install-ChocoPackage
{
    Param(
        [Parameter(Mandatory=$true)]
        [string]$PackageName,
        [string]$Version = ""
    )
    if ($Version -eq "")
    {
        choco upgrade $PackageName --yes
    }
    else
    {
        choco upgrade $PackageName --yes --version $Version
    }
}
using module "..\GetFileMetadata.psm1"

Param(
    [Parameter(Position = 0, Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ })]
    [ValidateScript({ (Get-Item $_).PSIsContainer })]
    [String] $Directory,

    [Parameter(Position = 1)]
    [String] $FilterString = "",

    [Parameter(Position = 2)]
    [switch] $IncludesSubdir,

    [Parameter(Position = 3)]
    [String[]] $PropertyNames=@()
)

$ErrorActionPreference = "Continue"
Set-StrictMode -Version 3.0

$params = @{
    Directory = $Directory
    FilterString = $FilterString
    IncludesSubdir = $IncludesSubdir
    PropertyNames = $PropertyNames
}

try {
    Get-FilesPropertiesInFolder @params
}
catch {
    Write-Error $_
}
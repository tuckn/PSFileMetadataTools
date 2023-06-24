using module "..\GetFileMetadata.psm1"

Param(
    [Parameter(Position = 0, Mandatory = $true)]
    [String] $FilePath,

    [Parameter(Position = 2)]
    [String[]] $PropertyNames=@()
)

$ErrorActionPreference = "Continue"
Set-StrictMode -Version 3.0

$params = @{
    FilePath = $FilePath
    PropertyNames = $PropertyNames
}

try {
    Get-FileProperties @params
}
catch {
    Write-Error $_
}
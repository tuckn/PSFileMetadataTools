using module "..\FileMetadataTools.psm1"

Param(
    [Parameter(Position = 0, Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ })]
    [ValidateScript({ (Get-Item $_).PSIsContainer })]
    [String] $Directory,

    [Parameter(Position = 1, Mandatory = $true)]
    [String] $ListFilePath,

    [Parameter(Position = 2)]
    [String] $FilterString = "",

    [Parameter(Position = 3)]
    [switch] $IncludesSubdir,

    [Parameter(Position = 4)]
    [String[]] $PropertyNames=@(),

    [Parameter(Position = 5)]
    [String] $ListFileEncoding = "utf8",

    [Parameter(Position = 6)]
    [switch] $Force,

    [Parameter(Position = 7)]
    [switch] $SmartOverwrite
)

$ErrorActionPreference = "Continue"
Set-StrictMode -Version 3.0

$params = @{
    Directory = $Directory
    ListFilePath = $ListFilePath
    FilterString = $FilterString
    IncludesSubdir = $IncludesSubdir
    PropertyNames = $PropertyNames
    ListFileEncoding = $ListFileEncoding
    Force = $Force
    SmartOverwrite = $SmartOverwrite
}

try {
    New-ListOfFilesProperties @params
}
catch {
    Write-Error $_
}
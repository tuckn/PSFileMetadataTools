using module "..\FileMetadataTools.psm1"

Param(
    [Parameter(Position = 0, Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ })]
    [String] $ExcelFilePath,

    [Parameter(Position = 1, Mandatory = $true)]
    [String[]] $SettingProperties=@()
)

$ErrorActionPreference = "Continue"
Set-StrictMode -Version 3.0

$params = @{
    ExcelFilePath = $ExcelFilePath
    SettingProperties = $SettingProperties
}

try {
    Set-ExcelProperties @params
}
catch {
    Write-Error $_
}
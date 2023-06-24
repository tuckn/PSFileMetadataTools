using module "..\GetFileMetadata.psm1"

Param(
    [Parameter(Position = 0, Mandatory = $true)]
    [String] $FilePath
)

$ErrorActionPreference = "Continue"
Set-StrictMode -Version 2.0

# FilePath is Foler
if ((Get-Item -LiteralPath $FilePath).PSIsContainer) {
    $childPath = Join-Path -Path "$FilePath" -ChildPath "*.*"
    foreach ($f in Get-ChildItem $childPath) {
        try {
            Get-FileProperties -FilePath "$($f.FullName)"
        }
        catch {
            Write-Error $_
        }
    }
}
# FilePath is File
elseif (Test-Path -LiteralPath $FilePath) {
    try {
        Get-FileProperties -FilePath "$FilePath"
    }
    catch {
        Write-Error $_
    }
}
else {
    Write-Error "The file is not existing: `"$FilePath`""
    exit 1
}

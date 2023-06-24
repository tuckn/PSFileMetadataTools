<#
.Synopsis
Obtains property information for the files contained in a given folder.

.Description
The properties to be retrieved are those that can be viewed in Windows Explorer.
Returns type [List[PSCustomObject]].

.Parameter Directory
A folder path

.Parameter FilterString
A path to the file to be obtain.

.Parameter IncludesSubdir
Whether files contained in sub folders are also included or not. If not specified, they will not be included.

.Parameter PropertyNames
[String[]] Property name to be obtain.
You can filter the properties to be retrieved. If nothing is specified, all properties will be returned.

.Example
PS> Get-ListOfFileProperties -FilePath "C:\Notes" FilterString ".xls*"
...
#>
using namespace System.Collections.Generic # PowerShell 5
$ErrorActionPreference = "Stop"
Set-StrictMode -Version 3.0

. (Join-Path -Path $PSScriptRoot -ChildPath "./Get-ListOfFileProperties.ps1")

function New-ListOfFilesProperties {
    [CmdletBinding()]
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
        [switch] $Force
    )
    Process {
        if (Test-Path -LiteralPath $ListFilePath) {
            if ($Force) {
                Write-Warning "[warn] `The existing list file will be overwrited: `"$($ListFilePath)`""
            }
            else {
                Write-Error "[error] `The list file is existing: `"$($ListFilePath)`". If you want to overwrite this, remove it or use -Force option."
            }
        }

        # Initializing the list of file properties in the folder
        [List[PSCustomObject]] $list = [List[PSCustomObject]]::new()

        $params = @{
            Directory = $Directory
            FilterString = $FilterString
            IncludesSubdir = $IncludesSubdir
            PropertyNames = $PropertyNames
        }

        try {
            $list = Get-ListOfFileProperties @params
        }
        catch {
            Write-Error $_
        }

        # Writing the list of file properties to the path
        try {
            ConvertTo-Json -InputObject $list | Out-File -LiteralPath $ListFilePath -Encoding $ListFileEncoding
        }
        catch {
            Write-Host "[error] $($_.Exception.Message)"
        }
    }
}

Export-ModuleMember -Function New-ListOfFilesProperties
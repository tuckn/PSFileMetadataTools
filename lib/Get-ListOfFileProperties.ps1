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

. (Join-Path -Path $PSScriptRoot -ChildPath "./Get-FileProperties.ps1")

function Get-ListOfFileProperties {
    [CmdletBinding()]
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
    Process {
        if (-Not (Get-Item -LiteralPath $Directory).PSIsContainer) {
            Write-Error "$($Directory) is not directory"
            exit 1
        }

        [String] $childPath = Join-Path -Path $Directory -ChildPath $FilterString

        [System.IO.FileInfo[]] $fileInfo = . {
            if ($IncludesSubdir) {
                return Get-ChildItem -Path $childPath -File -Recurse
            }
            else {
                return Get-ChildItem -Path $childPath -File
            }
        }

        $ls = [List[PSCustomObject]]::new()

        foreach ($f in $FileInfo) {
            $props = Get-FileProperties -FilePath $f.FullName -PropertyNames $PropertyNames
            $ls.Add($props)
        }

        return $ls
    }
}
Export-ModuleMember -Function Get-ListOfFileProperties
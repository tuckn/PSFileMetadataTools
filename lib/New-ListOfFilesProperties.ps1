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
        [switch] $Force,

        [Parameter(Position = 7)]
        [switch] $SmartOverwrite
    )
    Process {
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

        # Creating a tmp filepath
        $tmpFilePath = Join-Path -Path $env:TMP -ChildPath ([System.IO.Path]::GetRandomFileName())
        # Writing the list file to the tmp path
        try {
            ConvertTo-Json -InputObject $list | Out-File -LiteralPath $tmpFilePath -Encoding $ListFileEncoding
        }
        catch {
            Write-Error "[error] occured when writing a tmp file: $($tmpFilePath). Exception.Message: $($_.Exception.Message)"
        }

        if (Test-Path -LiteralPath $ListFilePath) {
            if ($SmartOverwrite) {
                Write-Warning "[warn] The listing file already exists: `"$($ListFilePath)`". If there are any changes to the contents, the file maight be overwritten."
                # Calculation MD5
                $md5Original = (Get-FileHash -Path $ListFilePath -Algorithm MD5).Hash
                $md5New = (Get-FileHash -Path $tmpFilePath -Algorithm MD5).Hash

                # Debuging
                Write-Host $md5Original
                Write-Host $md5New

                if ($md5Original -ne $md5New) {
                    Move-Item -Path $tmpFilePath -Destination $ListFilePath -Force
                }
                else {
                    Write-Host "[info] The list file was not updated because there were no changes to the file contents."
                }
            }
            elseif ($Force) {
                Write-Warning "[warn] The listing file already exists: `"$($ListFilePath)`". The file will be overwritten.."
                Move-Item -Path $tmpFilePath -Destination $ListFilePath -Force

            }
            else {
                Write-Error "[warn] The listing file already exists: `"$($ListFilePath)`". If you want to overwrite this, remove it or use -Force or -SmartOverwrite option."
            }
        }
        else {
            Move-Item -Path $tmpFilePath -Destination $ListFilePath -Force
        }
    }
}

Export-ModuleMember -Function New-ListOfFilesProperties
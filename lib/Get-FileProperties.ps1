<#
.Synopsis
Obtains properties for a specified file.

.Description
The properties that can be retrieved are the same as those displayed in Windows Explorer, including file name, size, creation date, modification date, access date, and attributes.

.Parameter FilePath
A path to the file to be obtain.

.Example
PS> Get-FileProperties -FilePath "C:\MyExcelNote.xlsx"
Name                        : MyExcelNote.xlsx
Size                        : 81.1 KB
Item type                   : Microsoft Excel Worksheet
Date modified               : 6/18/2023 9:28 AM
Date created                : 5/1/2023 3:32 PM
Date accessed               : 6/24/2023 6:03 AM
Attributes                  : ALP
Offline status              :
Availability                :
...
..

.Example
PS> Get-FileProperties -FilePath "C:\MyExcelNote.xlsx" | Set-Clipboard
#>
$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

function Get-FileProperties {
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0, Mandatory = $true)]
        [ValidateScript({ Test-Path -LiteralPath $_ })]
        [String] $FilePath
    )
    Process {
        Write-Host $FilePath

        $f = $null
        try {
            $f = Get-Item -LiteralPath "$FilePath"
        }
        catch {
            Write-Error $_
            exit 1
        }

        $sh = New-Object -COMObject Shell.Application
        $parentDir = Split-Path -Path $f
        $filename = Split-Path -Path $f -Leaf

        $shDir = $sh.Namespace($parentDir)
        $shFile = $shDir.ParseName($filename)

        [PSCustomObject] $props = New-Object -TypeName PSObject -Property @{}

        0..287 | ForEach-Object {
            if ($shDir.GetDetailsOf($null, $_)) {
                $propName = $shDir.GetDetailsOf($null, $_)
                $value = $shDir.GetDetailsOf($shFile, $_)
                # Write-Host "$($propName): $value" # Debug

                $propName = New-Object -TypeName PSNoteProperty -ArgumentList $propName, $value
                $props.PSObject.Properties.Add($propName)
            }
        }

        return $props
    }
}
Export-ModuleMember -Function Get-FileProperties
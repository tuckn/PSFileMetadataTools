<#
.Synopsis
Obtains properties for a specified file.

.Description
The properties that can be retrieved are the same as those displayed in Windows Explorer, including file name, size, creation date, modification date, access date, and attributes.
Returns type [PSCustomObject].

.Parameter FilePath
[String] A path to the file to be obtain.

.Parameter PropertyNames
[String[]] Property name to be obtain.
You can filter the properties to be retrieved. If nothing is specified, all properties will be returned.

.Example
Returns the same file information that Windows Explorer returns. However, modified datetime and created datetime return the universally coordinated datetime in ISO 8601 format. ex: 2016-02-07T09:03:47Z

PS> Get-FileProperties -FilePath "C:\MyExcelNote.xlsx"
Name                        : MyExcelNote.xlsx
Size                        : 81.1 KB
Item type                   : Microsoft Excel Worksheet
Date modified               : 2016-02-07T09:03:47Z
Date created                : 2023-05-01T14:53:38Z
Date accessed               : 6/24/2023 6:03 AM
Attributes                  : ALP
Offline status              :
Availability                :
...
..

.Example
You can filter the properties returned by the `PropertyNames` option.

PS> Get-FileProperties -FilePath "C:\MyExcelNote.xlsx" -PropertyNames "Name","Title","Categories"
Name                Title                    Categories
----                -----                    ----------
MyExcelNote.xlsx    My Excel Note 2023       Private; note

.Example
PS> Get-FileProperties -FilePath "C:\MyExcelNote.xlsx" | Set-Clipboard
#>
$ErrorActionPreference = "Stop"
Set-StrictMode -Version 3.0

function Get-FileProperties {
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0, Mandatory = $true)]
        [ValidateScript({ Test-Path -LiteralPath $_ })]
        [String] $FilePath,

        [Parameter(Position = 2)]
        [String[]] $PropertyNames=@()
    )
    Process {
        # Write-Host $FilePath # debug
        $f = $null
        try {
            $f = Get-Item -LiteralPath "$FilePath"
        }
        catch {
            Write-Error $_
            exit 1
        }

        $sh = New-Object -COMObject Shell.Application
        [String] $parentDir = Split-Path -Path $f
        [String] $filename = Split-Path -Path $f -Leaf

        [__ComObject] $shDir = $sh.Namespace($parentDir)
        [__ComObject] $shFile = $shDir.ParseName($filename)

        [PSCustomObject] $props = New-Object -TypeName PSObject -Property @{}
        [String] $propName = ""
        [String] $value = ""

        0..287 | ForEach-Object {
            if ($shDir.GetDetailsOf($null, $_)) {
                $propName = $shDir.GetDetailsOf($null, $_)

                if (
                    ($PropertyNames.Length -gt 0) -and
                    (-Not (($PropertyNames.Length -eq 1) -and
                        ([String]::IsNullOrEmpty($PropertyNames[0])))) -and
                    (-Not ($PropertyNames -icontains $propName))
                ) {
                    return
                }

                # @NOTE
                # Shell.Application returns datetime as a string, such as 'Data Created' or 'Data modified'. And it returns only up to the minute.
                # Example: Data modified: 6/18/2020 9:28 AM
                # So, use 'Get-Item' to get the datetime
                if ($propName -eq "Date created") {
                    # $value = $f.CreationTime.ToString("yyyy-MM-ddTHH:mm:sszzz")
                    $value = $f.CreationTime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
                }
                elseif ($propName -eq "Date modified") {
                    $value = $f.LastWriteTime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
                }
                else {
                    $value = $shDir.GetDetailsOf($shFile, $_)
                    # Write-Host "$($propName): $value" # Debug
                }

                $prop = New-Object -TypeName PSNoteProperty -ArgumentList $propName, $value
                $props.PSObject.Properties.Add($prop)
            }
        }

        return $props
    }
}
Export-ModuleMember -Function Get-FileProperties
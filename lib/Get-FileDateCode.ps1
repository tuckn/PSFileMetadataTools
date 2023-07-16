<#
.Synopsis
Obtains the date code for a specified file.

.Description
The older of the modification date and creation date of the specified file is adopted for the date information.

.Parameter FilePath
A path to the file to be obtain.

.Parameter DateFormat
The date format. For example "yyyy-MM-dd".
Default is "yyyy-MM-ddTHH:mm:sszzz".

.Example
PS> Get-FileDateCode -FilePath "C:\myphoto.jpg"
Created:  2018/11/15 19:44:01
Modefied: 2021/12/31 18:22:21
2018-11-15T19:44:01+09:00

.Example
PS> Get-FileDateCode -FilePath "C:\myphoto.jpg" -DateFormat "yy-MM-dd" | Set-Clipboard
Created:  2018/11/15 19:44:01
Modefied: 2021/12/31 18:22:21
#>
$ErrorActionPreference = "Stop"
Set-StrictMode -Version 3.0

function Get-FileDateCode {
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0, Mandatory = $true)]
        [ValidateScript({ Test-Path -LiteralPath $_ })]
        [String] $FilePath,

        [Parameter(Position = 1)]
        [String] $DateFormat = "yyyy-MM-ddTHH:mm:sszzz" # ISO 8601
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

        # Select the older date time
        Write-Host ('Created:  {0}' -f $f.CreationTime)
        Write-Host ('Modefied: {0}' -f $f.LastWriteTime)

        $d = $f.CreationTime
        if ($f.LastWriteTime -lt $f.CreationTime) {
            $d = $f.LastWriteTime
        }

        # @TODO: Get Meta data from EXIF, IPTC and so on...

        $dateCode = $d.ToString($DateFormat)

        return $dateCode
    }
}
Export-ModuleMember -Function Get-FileDateCode
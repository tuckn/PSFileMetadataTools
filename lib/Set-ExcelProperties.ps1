<#
.Synopsis
Obtains property information for the files contained in a given folder.

.Description
The properties to be retrieved are those that can be viewed in Windows Explorer.
Returns type [List[PSCustomObject]].

.Parameter ExcelFilePath
A Excel file path

.Parameter SettingProperties
A array of property string to set.
Application              Property   string Application {get;set;}
AppVersion               Property   string AppVersion {get;set;}
Author                   Property   string Author {get;set;}
Category                 Property   string Category {get;set;}
Comments                 Property   string Comments {get;set;}
Company                  Property   string Company {get;set;}
CorePropertiesXml        Property   xml CorePropertiesXml {get;}
Created                  Property   datetime Created {get;set;}
CustomPropertiesXml      Property   xml CustomPropertiesXml {get;}
ExtendedPropertiesXml    Property   xml ExtendedPropertiesXml {get;}
HyperlinkBase            Property   uri HyperlinkBase {get;set;}
HyperlinksChanged        Property   bool HyperlinksChanged {get;set;}
Keywords                 Property   string Keywords {get;set;}
LastModifiedBy           Property   string LastModifiedBy {get;set;}
LastPrinted              Property   string LastPrinted {get;set;}
LinksUpToDate            Property   bool LinksUpToDate {get;set;}
Manager                  Property   string Manager {get;set;}
Modified                 Property   datetime Modified {get;set;}
ScaleCrop                Property   bool ScaleCrop {get;set;}
SharedDoc                Property   bool SharedDoc {get;set;}
Status                   Property   string Status {get;set;}
Subject                  Property   string Subject {get;set;}
Title                    Property   string Title {get;set;}

.Example
PS> Get-ListOfFileProperties -FilePath "C:\Notes" FilterString ".xls*"
...
#>
using namespace System.Collections.Generic # PowerShell 5
$ErrorActionPreference = "Stop"
Set-StrictMode -Version 3.0

Add-Type -Path (Join-Path -Path $PSScriptRoot -ChildPath ".\EPPlus.dll")
. (Join-Path -Path $PSScriptRoot -ChildPath "./Get-ListOfFileProperties.ps1")

function Set-ExcelProperties {
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0, Mandatory = $true)]
        [ValidateScript({ Test-Path -LiteralPath $_ })]
        [String] $ExcelFilePath,

        [Parameter(Position = 1, Mandatory = $true)]
        [String[]] $SettingProperties=@()
    )
    Process {
        $xlFile = $null
        try {
            $xlFile = New-Object -TypeName OfficeOpenXml.ExcelPackage -ArgumentList $ExcelFilePath

            $wbProps = $xlFile.Workbook.Properties

            # $wbProps | Get-Member | Select-Object
            <#
Name                     MemberType Definition
----                     ---------- ----------
Equals                   Method     bool Equals(System.Object obj)
GetCustomPropertyValue   Method     System.Object GetCustomPropertyValue(string propertyName)
GetExtendedPropertyValue Method     string GetExtendedPropertyValue(string propertyName)
GetHashCode              Method     int GetHashCode()
GetType                  Method     type GetType()
SetCustomPropertyValue   Method     void SetCustomPropertyValue(string propertyName, System.Object value)
SetExtendedPropertyValue Method     void SetExtendedPropertyValue(string propertyName, string value)
ToString                 Method     string ToString()
Application              Property   string Application {get;set;}
AppVersion               Property   string AppVersion {get;set;}
Author                   Property   string Author {get;set;}
Category                 Property   string Category {get;set;}
Comments                 Property   string Comments {get;set;}
Company                  Property   string Company {get;set;}
CorePropertiesXml        Property   xml CorePropertiesXml {get;}
Created                  Property   datetime Created {get;set;}
CustomPropertiesXml      Property   xml CustomPropertiesXml {get;}
ExtendedPropertiesXml    Property   xml ExtendedPropertiesXml {get;}
HyperlinkBase            Property   uri HyperlinkBase {get;set;}
HyperlinksChanged        Property   bool HyperlinksChanged {get;set;}
Keywords                 Property   string Keywords {get;set;}
LastModifiedBy           Property   string LastModifiedBy {get;set;}
LastPrinted              Property   string LastPrinted {get;set;}
LinksUpToDate            Property   bool LinksUpToDate {get;set;}
Manager                  Property   string Manager {get;set;}
Modified                 Property   datetime Modified {get;set;}
ScaleCrop                Property   bool ScaleCrop {get;set;}
SharedDoc                Property   bool SharedDoc {get;set;}
Status                   Property   string Status {get;set;}
Subject                  Property   string Subject {get;set;}
Title                    Property   string Title {get;set;}
#>

            foreach ($setPropStr in $SettingProperties) {
                $parts = $setPropStr -split "=", 2
                $propName = $parts[0].Trim()
                $propValue = $parts[1].Trim()

                # Write-Host $propName
                # Write-Host $propValue
                $wbProps.$propName = $propValue
            }

            $xlFile.Save()
        }
        catch {
            Write-Error $_
        }
        finally {
            if ($null -ne $xlFile) {
                $xlFile.Dispose()
            }
        }
    }
}

Export-ModuleMember -Function Set-ExcelProperties
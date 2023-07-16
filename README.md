# PSFileMetadataTools

## How to use

### Get-FileDateCode

Returns the earliest of the file's creation date and modification date.

```powershell
PS> .\scripts\GetFileDateCode.ps1 "C:\myphoto.jpg"
Created:  2018/11/15 19:44:01
Modefied: 2021/12/31 18:22:21
2018-11-15T19:44:01+09:00
```

```powershell
PS> .\scripts\GetFileDateCode.ps1 "C:\myphoto.jpg" -DateFormat "yy-MM-dd"
Created:  2018/11/15 19:44:01
Modefied: 2021/12/31 18:22:21
2018-11-15
```

### Get-FileProperties

Returns the same file information that Windows Explorer returns. However, modified datetime and created datetime return the universally coordinated datetime in ISO 8601 format. ex: 2016-02-07T09:03:47Z

```powershell
PS> .\scripts\GetFileProperties.ps1 "C:\MyExcelNote.xlsx"
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
```

You can filter the properties returned by the `PropertyNames` option.

```powershell
PS> .\scripts\GetFileProperties.ps1 -FilePath "C:\MyExcelNote.xlsx" -PropertyNames "Name","Title","Categories"
Name                Title                    Categories
----                -----                    ----------
MyExcelNote.xlsx    My Excel Note 2023       Private; note
```

### New-ListOfFilesProperties

The following command will retrieve the properties of all files in the specified folder and save the results as `metadata.json`.

```powershell
PS> .\scripts\NewListOfFilesProperties.ps1 -Directory "C:\MyExcelNoteFolder" -PropertyNames "Name","Title","Tags" -ListFilePath "C:\MyExcelNoteFolder\.metadata.json" -SmartOverwrite
```

Since the `-SmartOverwrite` option is specified, if `metadata.json` already exists, it will be overwritten only when the contents are changed. In this case, the JSON file is not updated, thus reducing the resources of the script triggered by the update.

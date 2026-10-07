# SharePoint Bulk Upload Tool

A PowerShell tool that uploads many documents to a SharePoint Server (on-premises) document library at once, and fills in their metadata from a CSV file.

It is built for non-technical users: no script editing, guided prompts, and every check runs **before** anything is uploaded.

## Features

- Prompts for the site URL (with a default from `Config.psd1`)
- Opens a file picker to select the metadata CSV
- Lists the site's document libraries to choose from
- Lets you browse and pick the destination folder
- Matches CSV columns to library columns by display name or internal name
- Checks the CSV before uploading:
  - Missing, duplicate, or blank file names
  - Unknown column names (and lists the valid ones)
  - Required columns left blank
  - Invalid dates, numbers, Yes/No values, and choice values
- Shows a review screen and asks for confirmation before uploading
- Displays a progress bar during the upload
- Writes a log file (SUCCESS / FAILED / SKIPPED) for every document

## Requirements

- Windows PowerShell 5.1 (not PowerShell 7)
- SharePoint Server 2016, 2019, or Subscription Edition
- `SharePointPnPPowerShell2019` module:

```powershell
Install-Module SharePointPnPPowerShell2019 -Scope CurrentUser -AllowClobber
```

- Permission to add files to the target library

## Files

| File | Purpose |
| --- | --- |
| `SharePointBulkUploader.ps1` | The upload tool |
| `Upload.bat` | Double-click launcher for the tool |
| `Config.psd1` | Default site URL and upload delay |
| `metadata-sample.csv` | Example metadata file |

## Setup

1. Download the repository and keep all files in the same folder.
2. Edit `Config.psd1` and set `DefaultSiteUrl` to your site.
3. If Windows blocks the script, run:

```powershell
Unblock-File .\SharePointBulkUploader.ps1
```

## Preparing the CSV

- Put the CSV in the **same folder** as the documents you want to upload.
- The CSV must have a `FileName` column with the exact file name, including the extension.
- Every other column must match a column in the target library (display name or internal name).
- For multi-choice columns, separate values with a semicolon (`;`).

```csv
FileName,Title,DocumentDate
Report_001.docx,Monthly Report January,2025-01-15
Report_002.docx,Monthly Report February,2025-02-15
```

## Usage

**Easiest way:** double-click `Upload.bat`.

Or run it from PowerShell:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\SharePointBulkUploader.ps1
```

1. Enter the site URL, or press Enter to use the default.
2. Select your metadata CSV.
3. Sign in with `DOMAIN\username` and your password.
4. Choose the document library.
5. Browse to the destination folder and type `U` to upload there.
6. Review the summary and type `Y` to start.

The log file is saved in the same folder as the CSV.

## Notes

- Close the documents and the CSV before running. Files open in Word or Excel can fail to upload.
- Credentials are requested at runtime and held in memory only. Never store passwords in the script or config file.
- System fields such as Created By and Modified cannot be set.

## License

MIT. See [LICENSE](LICENSE).

## Author

Eleandro Girgis

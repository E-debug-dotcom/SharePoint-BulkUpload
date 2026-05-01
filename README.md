# SharePoint-BulkUpload

PowerShell script to bulk upload documents with metadata to a SharePoint on-premises document library using [PnP.PowerShell](https://pnp.github.io/powershell/).

## Features

- Reads file list and metadata from a CSV
- Uploads documents to a specified library and subfolder
- Applies metadata (e.g., project number, document date) on upload
- Displays a real-time progress bar during upload
- Throttle delay between uploads to avoid overwhelming the server
- Logs results (SUCCESS, FAILED, SKIPPED) to a CSV file

## Prerequisites

- [PnP.PowerShell](https://pnp.github.io/powershell/) module

  ```powershell
  Install-Module PnP.PowerShell -Scope CurrentUser
  ```

- SharePoint on-premises (2016 / 2019 / SE) with appropriate permissions
- PowerShell execution policy set to `RemoteSigned` or higher

  ```powershell
  Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
  ```

## CSV Format

Your `metadata.csv` must include a `FileName` column with the exact filename (including extension). Additional columns map to SharePoint column **internal names**.

```csv
FileName,ProjectNum,DocumentDate
Report_001.docx,PROJ-A,2024-01-15
Report_002.docx,PROJ-B,2025-01-16
```

A sample file is included: [`metadata-sample.csv`](metadata-sample.csv)

## Configuration

Edit the `CONFIG` section at the top of `Upload-ToSharePoint.ps1`:

| Variable | Description | Example |
|---|---|---|
| `$SiteUrl` | SharePoint site collection URL | `https://yoursite.sharepoint.com/sites/YourSite` |
| `$LibraryName` | Target document library name | `Shared Documents` |
| `$TargetFolder` | Library-relative path to the target folder | `Shared Documents/subfolder` |
| `$CsvPath` | Full path to your metadata CSV | `C:\path\to\metadata.csv` |
| `$SourceFolder` | Folder containing the files to upload | `C:\path\to\source\files` |
| `$LogPath` | Full path for the output log CSV | `C:\path\to\upload-log.csv` |
| `$ThrottleMs` | Delay (ms) between uploads to avoid throttling | `300` |

## Usage

1. Edit the `CONFIG` section with your environment details
2. Place your source files in the `$SourceFolder` directory
3. Prepare your `metadata.csv` with filenames and metadata columns
4. Open PowerShell and run:

   ```powershell
   .\Upload-ToSharePoint.ps1
   ```

5. Enter your credentials when prompted
6. Monitor the progress bar and console output
7. Review `upload-log.csv` for detailed results

## Notes

- **Verify internal column names** — SharePoint internal names often differ from display names. Check via:

  ```powershell
  Connect-PnPOnline -Url "https://yoursite.sharepoint.com/sites/YourSite" -Credentials (Get-Credential)
  Get-PnPField -List "Your Library" | Select-Object Title, InternalName
  ```

  Use the `InternalName` values in the `$values` hashtable in the script.

- **Close files before running** — If a source file is open in Word or locked by Explorer's preview pane, the upload will fail for that file. Close all source files and toggle off the preview pane (`Alt+P`) before running.

- **Unblock downloaded scripts** — If you downloaded this script from the internet, Windows may block execution. Run:

  ```powershell
  Unblock-File -Path ".\Upload-ToSharePoint.ps1"
  ```

- **Read-only system fields** — Fields like `Modified By`, `Created By`, and `Modified` are system-managed. Do not include them in the `$values` hashtable.

## Author

**Eleandro Girgis**

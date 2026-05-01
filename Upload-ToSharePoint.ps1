# ============================================================
# Upload-ToSharePoint.ps1
# Bulk uploads documents to SharePoint on-prem with metadata.
# Reads from a CSV file and logs results per file.
#
# Author:  Eleandro Girgis
# Created: 2026-04-30
#
# Requirements:
#   Install-Module PnP.PowerShell -Scope CurrentUser
# ============================================================

# --- CONFIG -------------------------------------------------
$SiteUrl      = "https://###/####" # Full site URL
$LibraryName  = "Test_Bulk upload" # Target document library
$TargetFolder = "Test_Bulk upload/test" # Target folder path inside library
$CsvPath      = "C:\$UserName$\Downloads\metadata.csv" # Path to your metadata CSV
$SourceFolder = "C:\$UserName$\Downloads\bulkupload" # Folder containing the files
$LogPath      = "C:\$UserName$\Downloads\bulkupload\upload-log.csv" # Output log path
$ThrottleMs   = 300 # Delay between uploads (ms) to avoid throttling
# ------------------------------------------------------------

# Connect using credentials prompt
Connect-PnPOnline -Url $SiteUrl -Credentials (Get-Credential)

# Initialize log
$logEntries = [System.Collections.Generic.List[PSCustomObject]]::new()

# Import CSV
$metadata = Import-Csv -Path $CsvPath
$totalFiles = $metadata.Count
$currentFile = 0

Write-Host "Starting upload of $totalFiles documents..." -ForegroundColor Cyan

foreach ($row in $metadata) {

    $currentFile++
    $fileName = $row.FileName
    $percentComplete = [Math]::Round(($currentFile / $totalFiles) * 100)

    # Progress bar
    Write-Progress -Activity "Uploading documents to SharePoint" -Status "$currentFile of $totalFiles - $fileName" -PercentComplete $percentComplete

    $filePath = Join-Path $SourceFolder $fileName

    # Skip if file doesn't exist on disk
    if (-not (Test-Path $filePath)) {
        Write-Warning "File not found, skipping: $fileName"
        $logEntries.Add([PSCustomObject]@{
            FileName  = $fileName
            Status    = "SKIPPED - file not found"
            Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            Error     = ""
        })
        continue
    }

    try {
        # MAP CSV COLUMNS TO SHAREPOINT INTERNAL NAMES
        $values = @{
            #"ProjectNum"    = $row.ProjectNum
            "DocumentDate"  = $row.DocumentDate
        }

        # Upload file and set metadata
        Add-PnPFile -Path $filePath -Folder $TargetFolder -Values $values -ErrorAction Stop | Out-Null

        Write-Host "  OK  $fileName" -ForegroundColor Green

        $logEntries.Add([PSCustomObject]@{
            FileName  = $fileName
            Status    = "SUCCESS"
            Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            Error     = ""
        })
    }
    catch {
        Write-Warning "FAILED: $fileName -- $_"

        $logEntries.Add([PSCustomObject]@{
            FileName  = $fileName
            Status    = "FAILED"
            Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            Error     = $_.Exception.Message
        })
    }

    Start-Sleep -Milliseconds $ThrottleMs
}

# Clear the progress bar
Write-Progress -Activity "Uploading documents to SharePoint" -Completed

# Write log
$logEntries | Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8

# Summary

$succeeded = @($logEntries | Where-Object Status -eq "SUCCESS").Count
$failed    = @($logEntries | Where-Object Status -eq "FAILED").Count
$skipped   = @($logEntries | Where-Object Status -like "SKIPPED*").Count


Write-Host ""
Write-Host "--- Upload Complete ---" -ForegroundColor Cyan
Write-Host "  Succeeded : $succeeded"
Write-Host "  Failed    : $failed"
Write-Host "  Skipped   : $skipped"
Write-Host "  Log saved : $LogPath"

Disconnect-PnPOnline

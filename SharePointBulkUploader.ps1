<#
.SYNOPSIS
    Bulk upload documents and metadata to SharePoint (on-premises).

.DESCRIPTION
    Connects to a SharePoint site, lets the user pick a document library
    and destination folder, validates metadata from a CSV file against the
    library's columns, uploads the documents, and writes a detailed log.

    Nothing is uploaded until every check passes and the user confirms.

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\SharePointBulkUploader.ps1

.NOTES
    Version : 1.0.0
    Author  : Eleandro Girgis
    License : MIT

    Requirements:
      - Windows PowerShell 5.1 (not PowerShell 7)
      - SharePointPnPPowerShell2019 module
          Install-Module SharePointPnPPowerShell2019 -Scope CurrentUser -AllowClobber
      - SharePoint Server 2016 / 2019 / Subscription Edition
      - Config.psd1 in the same folder as this script
#>

$ScriptVersion = "1.0.0"

# --- Environment checks (run before asking the user anything) ---
if ($PSVersionTable.PSVersion.Major -ge 6) {
    Write-Host ""
    Write-Host "ERROR: Wrong version of PowerShell." -ForegroundColor Red
    Write-Host "This tool needs Windows PowerShell 5.1." -ForegroundColor Yellow
    Write-Host "You are using PowerShell $($PSVersionTable.PSVersion)." -ForegroundColor Yellow
    Write-Host "Run it with powershell.exe instead of pwsh.exe." -ForegroundColor Yellow
    exit 1
}

if (-not (Get-Module -ListAvailable -Name SharePointPnPPowerShell2019)) {
    Write-Host ""
    Write-Host "ERROR: A required module is not installed:" -ForegroundColor Red
    Write-Host "       SharePointPnPPowerShell2019" -ForegroundColor Yellow
    Write-Host "Install it with:" -ForegroundColor Yellow
    Write-Host "       Install-Module SharePointPnPPowerShell2019 -Scope CurrentUser -AllowClobber" -ForegroundColor Yellow
    exit 1
}

Remove-Module PnP.PowerShell -ErrorAction SilentlyContinue
try {
    Import-Module SharePointPnPPowerShell2019 -DisableNameChecking -WarningAction SilentlyContinue -ErrorAction Stop
}
catch {
    Write-Host ""
    Write-Host "ERROR: The SharePoint module could not be loaded." -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Yellow
    Write-Host "No documents were uploaded." -ForegroundColor Yellow
    exit 1
}

# --- Startup banner -------------------------------------------
Write-Host ""
Write-Host "=======================================================" -ForegroundColor Cyan
Write-Host " SHAREPOINT BULK UPLOAD TOOL v$ScriptVersion" -ForegroundColor Cyan
Write-Host "=======================================================" -ForegroundColor Cyan

# --- Load settings from Config.psd1 --------------------------
$ConfigPath = Join-Path $PSScriptRoot "Config.psd1"

if (-not (Test-Path -LiteralPath $ConfigPath)) {
    Write-Host ""
    Write-Host "ERROR: Settings file not found:" -ForegroundColor Red
    Write-Host "       $ConfigPath" -ForegroundColor Yellow
    Write-Host "Make sure Config.psd1 is in the same folder as this script." -ForegroundColor Yellow
    exit 1
}

try {
    $Config = Import-PowerShellDataFile -Path $ConfigPath -ErrorAction Stop
}
catch {
    Write-Host ""
    Write-Host "ERROR: Config.psd1 could not be read. Check it for typos (missing quotes or braces)." -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Yellow
    exit 1
}

$DefaultSiteUrl = "$($Config.DefaultSiteUrl)".Trim()
$ThrottleMs     = 300
if ($Config.ThrottleMs) { $ThrottleMs = [int]$Config.ThrottleMs }

# --- Ask which SharePoint site to upload to ------------------
do {
    Write-Host ""
    Write-Host "Which SharePoint site are you uploading to?" -ForegroundColor Cyan
    Write-Host "Paste the site URL only (not a library or page link)." -ForegroundColor DarkGray
    if ($DefaultSiteUrl) {
        Write-Host "Press Enter to use: $DefaultSiteUrl" -ForegroundColor DarkGray
    }

    $SiteUrl = "$(Read-Host 'Site URL')".Trim()
    if (-not $SiteUrl) { $SiteUrl = $DefaultSiteUrl }
    $SiteUrl = $SiteUrl.TrimEnd('/')

    $validSite = $true
    if ($SiteUrl -match '^https://https://') {
        Write-Host "The URL contains https:// twice. Remove one and try again." -ForegroundColor Red
        $validSite = $false
    }
    elseif ($SiteUrl -notmatch '^https://[^/\s]+(/\S*)?$') {
        Write-Host "That is not a valid URL. It must start with https://" -ForegroundColor Red
        $validSite = $false
    }
    elseif ($SiteUrl -match '/Forms/|/_layouts/|/SitePages/|\.aspx') {
        Write-Host "That looks like a library or page link. Enter the site URL instead." -ForegroundColor Red
        $validSite = $false
    }
} until ($validSite)

Write-Host "Site: $SiteUrl" -ForegroundColor Green

# --- Choose the metadata CSV ---------------------------------
Add-Type -AssemblyName System.Windows.Forms

Write-Host ""
Write-Host "Select the metadata CSV file in the window that opens..." -ForegroundColor Cyan
Write-Host "The documents must be in the same folder as the CSV." -ForegroundColor DarkGray

# Hidden "always on top" window, so the Open window shows in front
$owner = New-Object System.Windows.Forms.Form
$owner.TopMost = $true
$owner.ShowInTaskbar = $false

$dialog = New-Object System.Windows.Forms.OpenFileDialog
$dialog.Title = "Select the metadata CSV file"
$dialog.Filter = "CSV files (*.csv)|*.csv"
$dialog.Multiselect = $false
$dialog.InitialDirectory = "$env:USERPROFILE\Downloads"

$result = $dialog.ShowDialog($owner)
$owner.Dispose()

if ("$result" -ne "OK") {
    Write-Host ""
    Write-Host "No CSV selected. Nothing was uploaded." -ForegroundColor Yellow
    exit 0
}

$CsvPath      = $dialog.FileName
$SourceFolder = Split-Path $CsvPath -Parent
$timestamp    = Get-Date -Format "yyyy-MM-dd_HHmmss"
$LogPath      = Join-Path $SourceFolder "SharePoint_Upload_Log_$timestamp.csv"

Write-Host "CSV:           $CsvPath" -ForegroundColor Green
Write-Host "Source folder: $SourceFolder" -ForegroundColor Green
Write-Host "Log file:      $LogPath" -ForegroundColor DarkGray

# --- Check the CSV and files (before sign-in) ----------------
Write-Host ""
Write-Host "Checking the CSV and files..." -ForegroundColor Cyan

try {
    $metadata = @(Import-Csv -LiteralPath $CsvPath -ErrorAction Stop)
}
catch {
    Write-Host "ERROR: Could not open the CSV. If it is open in Excel, close it and try again." -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Yellow
    exit 1
}

if ($metadata.Count -eq 0) {
    Write-Host "ERROR: The CSV has no rows. Nothing was uploaded." -ForegroundColor Red
    exit 1
}

$csvColumns = @($metadata[0].PSObject.Properties.Name)
if ($csvColumns -notcontains "FileName") {
    Write-Host "ERROR: The CSV must have a column named FileName." -ForegroundColor Red
    Write-Host "Columns found: $($csvColumns -join ', ')" -ForegroundColor Yellow
    Write-Host "Nothing was uploaded." -ForegroundColor Yellow
    exit 1
}

# Every column except FileName is treated as a SharePoint column (matched below)
$metadataColumns = @($csvColumns | Where-Object { $_ -ne "FileName" })

$problems  = New-Object System.Collections.Generic.List[string]
$seen      = @{}
$rowNumber = 1   # Row 1 is the header, so the first document is row 2 (same as Excel)

foreach ($row in $metadata) {
    $rowNumber++
    $name = "$($row.FileName)".Trim()

    if (-not $name) {
        $problems.Add("Row ${rowNumber}: FileName is blank")
        continue
    }
    if ($name.Contains('\') -or $name.Contains('/')) {
        $problems.Add("Row ${rowNumber}: '$name' must be a file name only, not a path")
        continue
    }
    if ($seen.ContainsKey($name)) {
        $problems.Add("Row ${rowNumber}: '$name' is listed twice (also row $($seen[$name]))")
    }
    else {
        $seen[$name] = $rowNumber
    }
    if (-not (Test-Path -LiteralPath (Join-Path $SourceFolder $name) -PathType Leaf)) {
        $problems.Add("Row ${rowNumber}: '$name' was not found in the source folder")
    }
}

if ($problems.Count -gt 0) {
    Write-Host ""
    Write-Host "Found $($problems.Count) problem(s) in the CSV:" -ForegroundColor Red
    $problems | Select-Object -First 25 | ForEach-Object { Write-Host "  - $_" -ForegroundColor Yellow }
    if ($problems.Count -gt 25) {
        Write-Host "  ...and $($problems.Count - 25) more." -ForegroundColor Yellow
    }
    Write-Host ""
    Write-Host "Fix the CSV and run the tool again. Nothing was uploaded." -ForegroundColor Yellow
    exit 1
}

Write-Host "File check passed: $($metadata.Count) documents found." -ForegroundColor Green
if ($metadataColumns.Count -eq 0) {
    Write-Host "No metadata columns in the CSV. Files will upload without metadata." -ForegroundColor Yellow
}
else {
    Write-Host "Metadata columns in the CSV: $($metadataColumns -join ', ')" -ForegroundColor DarkGray
}

# --- Connect --------------------------------------------------
# Credentials are requested at runtime and held in memory only.
# Never hardcode usernames or passwords in this script.
Write-Host ""
Write-Host "Enter your SharePoint credentials." -ForegroundColor Cyan
$username = Read-Host "Username (DOMAIN\username)"
$password = Read-Host "Password" -AsSecureString
$cred = New-Object System.Management.Automation.PSCredential ($username, $password)

try {
    Connect-PnPOnline -Url $SiteUrl -Credentials $cred -ErrorAction Stop
}
catch {
    Write-Host ""
    Write-Host "ERROR: Could not connect to $SiteUrl" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Yellow
    Write-Host "Check the site URL, username, and password. Nothing was uploaded." -ForegroundColor Yellow
    exit 1
}

# --- Pre-flight check: verify connection BEFORE looping files ---
try {
    $web = Get-PnPWeb -ErrorAction Stop
    Write-Host "Connected to SharePoint: $($web.Title) ($($web.ServerRelativeUrl))" -ForegroundColor Green
}
catch {
    Write-Host ""
    Write-Host "ERROR: Connected, but could not read site info." -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Yellow
    Disconnect-PnPOnline
    exit 1
}

# Initialize log
$logEntries = New-Object System.Collections.Generic.List[PSCustomObject]

# --- Choose the document library -----------------------------
Write-Host ""
Write-Host "Loading document libraries..." -ForegroundColor Cyan

$libraries = @(Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 -and -not $_.Hidden } | Sort-Object Title)

if ($libraries.Count -eq 0) {
    Write-Host "ERROR: No document libraries found on this site, or you do not have access." -ForegroundColor Red
    Disconnect-PnPOnline
    exit 1
}

Write-Host ""
Write-Host "Document libraries on this site:" -ForegroundColor Cyan
for ($i = 0; $i -lt $libraries.Count; $i++) {
    Write-Host ("  {0,3}. {1}" -f ($i + 1), $libraries[$i].Title)
}

$prompt = "Type the number of the library (1-$($libraries.Count))"
do {
    Write-Host ""
    $choice = "$(Read-Host $prompt)".Trim()
    $validChoice = ($choice -match '^\d+$') -and ([int]$choice -ge 1) -and ([int]$choice -le $libraries.Count)
    if (-not $validChoice) {
        Write-Host "Please type a number from 1 to $($libraries.Count)." -ForegroundColor Red
    }
} until ($validChoice)

$list        = $libraries[[int]$choice - 1]
$LibraryName = $list.Title

# Work out the library's real folder path (it can differ from its display name)
$rootFolder   = Get-PnPProperty -ClientObject $list -Property RootFolder
$webPath      = "$($web.ServerRelativeUrl)".TrimEnd('/')
$TargetFolder = $rootFolder.ServerRelativeUrl.Substring($webPath.Length).TrimStart('/')

Write-Host "Library: $LibraryName" -ForegroundColor Green

# --- Load the library's columns (fields) ----------------------
Write-Host ""
Write-Host "Loading column settings for '$LibraryName'..." -ForegroundColor Cyan

$systemFields = @(
    'ContentType','Attachments','Edit','LinkTitle','LinkTitleNoMenu','DocIcon',
    'ItemChildCount','FolderChildCount','_ComplianceFlags','_ComplianceTag',
    '_ComplianceTagWrittenTime','_ComplianceTagUserId'
)

$fields = @(Get-PnPField -List $list -ErrorAction Stop | Where-Object {
    -not $_.Hidden -and -not $_.ReadOnlyField -and ($systemFields -notcontains $_.InternalName)
})

$fieldLookup = @{}
foreach ($field in $fields) {
    $fieldLookup[$field.Title.ToLower()]        = $field
    $fieldLookup[$field.InternalName.ToLower()] = $field
}

Write-Host "Found $($fields.Count) usable column(s) in this library." -ForegroundColor Green

# --- Match the CSV's metadata columns to the library's columns ---
$columnMap     = @{}   # CSV header -> field object
$matchProblems = New-Object System.Collections.Generic.List[string]

foreach ($csvHeader in $metadataColumns) {
    $key = $csvHeader.ToLower().Trim()
    if ($fieldLookup.ContainsKey($key)) {
        $columnMap[$csvHeader] = $fieldLookup[$key]
    }
    else {
        $matchProblems.Add("CSV column '$csvHeader' does not match any column in '$LibraryName'")
    }
}

if ($matchProblems.Count -gt 0) {
    Write-Host ""
    Write-Host "Found $($matchProblems.Count) column mismatch(es):" -ForegroundColor Red
    $matchProblems | ForEach-Object { Write-Host "  - $_" -ForegroundColor Yellow }

    Write-Host ""
    Write-Host "Columns available in '$LibraryName':" -ForegroundColor Cyan
    $fields | Sort-Object Title | ForEach-Object {
        Write-Host ("  - {0}  (internal name: {1}, type: {2})" -f $_.Title, $_.InternalName, $_.TypeAsString)
    }

    Write-Host ""
    Write-Host "Fix the CSV column headers and run the tool again. Nothing was uploaded." -ForegroundColor Yellow
    Disconnect-PnPOnline
    exit 1
}

if ($columnMap.Count -eq 0) {
    Write-Host "No metadata columns to match. Files will upload without metadata." -ForegroundColor Yellow
}
else {
    Write-Host "Matched $($columnMap.Count) CSV column(s) to library columns:" -ForegroundColor Green
    foreach ($csvHeader in $columnMap.Keys) {
        Write-Host ("  '{0}' -> {1} ({2})" -f $csvHeader, $columnMap[$csvHeader].Title, $columnMap[$csvHeader].TypeAsString) -ForegroundColor DarkGray
    }
}

# --- Check each CSV value against its column's type -----------
Write-Host ""
Write-Host "Checking metadata values..." -ForegroundColor Cyan

$valueProblems = New-Object System.Collections.Generic.List[string]
$choiceCache   = @{}
$rowNumber     = 1

foreach ($row in $metadata) {
    $rowNumber++

    foreach ($csvHeader in $columnMap.Keys) {
        $field = $columnMap[$csvHeader]
        $value = "$($row.$csvHeader)".Trim()

        if (-not $value) {
            if ($field.Required) {
                $valueProblems.Add("Row ${rowNumber}: '$($field.Title)' is required and is blank")
            }
            continue
        }

        switch ($field.TypeAsString) {
            'DateTime' {
                $parsedDate = [datetime]::MinValue
                if (-not [datetime]::TryParse($value, [ref]$parsedDate)) {
                    $valueProblems.Add("Row ${rowNumber}: '$($field.Title)' value '$value' is not a valid date")
                }
            }
            { $_ -in @('Number', 'Currency') } {
                $parsedNumber = 0.0
                if (-not [double]::TryParse($value, [ref]$parsedNumber)) {
                    $valueProblems.Add("Row ${rowNumber}: '$($field.Title)' value '$value' is not a valid number")
                }
            }
            'Boolean' {
                if ($value -notmatch '^(yes|no|true|false|1|0)$') {
                    $valueProblems.Add("Row ${rowNumber}: '$($field.Title)' value '$value' must be Yes or No")
                }
            }
            { $_ -in @('Choice', 'MultiChoice') } {
                if (-not $choiceCache.ContainsKey($field.InternalName)) {
                    $choiceField = Get-PnPProperty -ClientObject $field -Property Choices
                    $choiceCache[$field.InternalName] = @($choiceField.Choices)
                }
                $validChoices  = $choiceCache[$field.InternalName]
                $valuesToCheck = if ($field.TypeAsString -eq 'MultiChoice') { $value -split ';' } else { @($value) }
                foreach ($v in $valuesToCheck) {
                    $v = $v.Trim()
                    if ($v -and ($validChoices -notcontains $v)) {
                        $valueProblems.Add("Row ${rowNumber}: '$($field.Title)' value '$v' is not one of: $($validChoices -join ', ')")
                    }
                }
            }
            default {
                # Text, Note, and other column types are passed through as-is
            }
        }
    }
}

if ($valueProblems.Count -gt 0) {
    Write-Host ""
    Write-Host "Found $($valueProblems.Count) metadata value problem(s):" -ForegroundColor Red
    $valueProblems | Select-Object -First 25 | ForEach-Object { Write-Host "  - $_" -ForegroundColor Yellow }
    if ($valueProblems.Count -gt 25) {
        Write-Host "  ...and $($valueProblems.Count - 25) more." -ForegroundColor Yellow
    }
    Write-Host ""
    Write-Host "Fix the CSV and run the tool again. Nothing was uploaded." -ForegroundColor Yellow
    Disconnect-PnPOnline
    exit 1
}

Write-Host "Metadata check passed." -ForegroundColor Green

# --- Choose the destination folder ---------------------------
$libraryRoot   = $TargetFolder
$currentFolder = $libraryRoot

do {
    try {
        $subFolders = @(Get-PnPFolderItem -FolderSiteRelativeUrl $currentFolder -ItemType Folder -ErrorAction Stop |
            Where-Object { -not ($currentFolder -eq $libraryRoot -and $_.Name -eq 'Forms') } |
            Sort-Object Name)
    }
    catch {
        Write-Host "ERROR: Could not read folders in '$currentFolder'." -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Yellow
        Disconnect-PnPOnline
        exit 1
    }

    Write-Host ""
    Write-Host "Current folder: $currentFolder" -ForegroundColor Cyan
    if ($subFolders.Count -eq 0) {
        Write-Host "  (no subfolders)" -ForegroundColor DarkGray
    }
    for ($i = 0; $i -lt $subFolders.Count; $i++) {
        Write-Host ("  {0,3}. {1}" -f ($i + 1), $subFolders[$i].Name)
    }

    Write-Host ""
    if ($subFolders.Count -gt 0) {
        Write-Host "  Type a number to open that folder" -ForegroundColor DarkGray
    }
    Write-Host "  U = upload to the current folder" -ForegroundColor DarkGray
    if ($currentFolder -ne $libraryRoot) {
        Write-Host "  B = go back up one level" -ForegroundColor DarkGray
    }

    $choice = "$(Read-Host 'Your choice')".Trim().ToUpper()
    $done   = $false

    if ($choice -eq 'U') {
        $done = $true
    }
    elseif ($choice -eq 'B' -and $currentFolder -ne $libraryRoot) {
        $currentFolder = $currentFolder.Substring(0, $currentFolder.LastIndexOf('/'))
    }
    elseif ($choice -match '^\d+$' -and [int]$choice -ge 1 -and [int]$choice -le $subFolders.Count) {
        $currentFolder = "$currentFolder/$($subFolders[[int]$choice - 1].Name)"
    }
    else {
        Write-Host "Not a valid choice. Try again." -ForegroundColor Red
    }
} until ($done)

$TargetFolder = $currentFolder
Write-Host "Upload folder: $TargetFolder" -ForegroundColor Green

# --- Review before uploading -----------------------------------
Write-Host ""
Write-Host "=======================================================" -ForegroundColor Cyan
Write-Host "REVIEW UPLOAD" -ForegroundColor Cyan
Write-Host "=======================================================" -ForegroundColor Cyan
Write-Host "Site:             $SiteUrl"
Write-Host "Library:          $LibraryName"
Write-Host "Upload folder:    $TargetFolder"
Write-Host "CSV:              $CsvPath"
Write-Host "Documents:        $($metadata.Count)"
if ($columnMap.Count -eq 0) {
    Write-Host "Metadata columns: (none)"
}
else {
    Write-Host "Metadata columns: $($columnMap.Keys -join ', ')"
}
Write-Host "Log file:         $LogPath"
Write-Host "=======================================================" -ForegroundColor Cyan

do {
    $confirm = "$(Read-Host 'Continue with upload? (Y/N)')".Trim().ToUpper()
} until ($confirm -eq 'Y' -or $confirm -eq 'N')

if ($confirm -eq 'N') {
    Write-Host ""
    Write-Host "Upload cancelled. Nothing was uploaded." -ForegroundColor Yellow
    Disconnect-PnPOnline
    exit 0
}

Write-Host "Starting upload of $($metadata.Count) documents..." -ForegroundColor Cyan

$totalCount   = $metadata.Count
$currentCount = 0

foreach ($row in $metadata) {

    $currentCount++
    $fileName = "$($row.FileName)".Trim()
    $filePath = Join-Path $SourceFolder $fileName

    $percentComplete = [math]::Round(($currentCount / $totalCount) * 100)
    Write-Progress -Activity "Uploading documents to SharePoint" `
        -Status "$currentCount of $totalCount - $fileName" `
        -PercentComplete $percentComplete

    $startTime = Get-Date

    if (-not (Test-Path -LiteralPath $filePath -PathType Leaf)) {
        Write-Warning "File not found, skipping: $fileName"
        $endTime = Get-Date
        $logEntries.Add([PSCustomObject]@{
            FileName        = $fileName
            Status          = "SKIPPED - file not found"
            SourcePath      = $filePath
            Destination     = $TargetFolder
            StartTime       = $startTime.ToString("yyyy-MM-dd HH:mm:ss")
            EndTime         = $endTime.ToString("yyyy-MM-dd HH:mm:ss")
            DurationSeconds = [math]::Round(($endTime - $startTime).TotalSeconds, 2)
            Error           = ""
        })
        continue
    }

    try {
        # Build metadata values from the matched library columns,
        # not hardcoded column names, so this works for any library.
        $values = @{}
        foreach ($csvHeader in $columnMap.Keys) {
            $values[$columnMap[$csvHeader].InternalName] = $row.$csvHeader
        }

        Add-PnPFile `
            -Path       $filePath `
            -Folder     $TargetFolder `
            -Values     $values `
            -ErrorAction Stop | Out-Null

        $endTime = Get-Date
        Write-Host "  OK  $fileName" -ForegroundColor Green

        $logEntries.Add([PSCustomObject]@{
            FileName        = $fileName
            Status          = "SUCCESS"
            SourcePath      = $filePath
            Destination     = $TargetFolder
            StartTime       = $startTime.ToString("yyyy-MM-dd HH:mm:ss")
            EndTime         = $endTime.ToString("yyyy-MM-dd HH:mm:ss")
            DurationSeconds = [math]::Round(($endTime - $startTime).TotalSeconds, 2)
            Error           = ""
        })
    }
    catch {
        $endTime = Get-Date
        Write-Warning "FAILED: $fileName -- $($_.Exception.Message)"

        $logEntries.Add([PSCustomObject]@{
            FileName        = $fileName
            Status          = "FAILED"
            SourcePath      = $filePath
            Destination     = $TargetFolder
            StartTime       = $startTime.ToString("yyyy-MM-dd HH:mm:ss")
            EndTime         = $endTime.ToString("yyyy-MM-dd HH:mm:ss")
            DurationSeconds = [math]::Round(($endTime - $startTime).TotalSeconds, 2)
            Error           = $_.Exception.Message
        })
    }

    Start-Sleep -Milliseconds $ThrottleMs
}

Write-Progress -Activity "Uploading documents to SharePoint" -Completed

$logEntries | Export-Csv -LiteralPath $LogPath -NoTypeInformation -Encoding UTF8

# @() makes .Count reliable in PowerShell 5.1 when only one entry matches
$succeeded = @($logEntries | Where-Object Status -eq "SUCCESS").Count
$failed    = @($logEntries | Where-Object Status -eq "FAILED").Count
$skipped   = @($logEntries | Where-Object Status -like "SKIPPED*").Count

Write-Host ""
Write-Host "=======================================================" -ForegroundColor Cyan
Write-Host "UPLOAD COMPLETE" -ForegroundColor Cyan
Write-Host "=======================================================" -ForegroundColor Cyan
Write-Host ("  Succeeded : {0}" -f $succeeded) -ForegroundColor Green
if ($failed -gt 0) {
    Write-Host ("  Failed    : {0}" -f $failed) -ForegroundColor Red
}
else {
    Write-Host ("  Failed    : {0}" -f $failed)
}
if ($skipped -gt 0) {
    Write-Host ("  Skipped   : {0}" -f $skipped) -ForegroundColor Yellow
}
else {
    Write-Host ("  Skipped   : {0}" -f $skipped)
}
Write-Host ""
Write-Host "  Log file:"
Write-Host "  $LogPath" -ForegroundColor DarkGray
Write-Host "=======================================================" -ForegroundColor Cyan

if ($failed -gt 0 -or $skipped -gt 0) {
    Write-Host ""
    Write-Host "Some documents were not uploaded successfully." -ForegroundColor Yellow
    Write-Host "Open the log file above to see which ones, and why." -ForegroundColor Yellow
}

Disconnect-PnPOnline

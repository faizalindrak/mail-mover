<#
.SYNOPSIS
    Imports EML files into Outlook PST using COM automation.

.DESCRIPTION
    Reads .eml files from a folder structure, parses them via CDO, and injects
    directly into Outlook PST folders using PostItem + MessageClass swap trick.

    Classic Outlook MUST be running during execution.

.NOTES
    Requires: Classic Outlook (not New Outlook), PowerShell
    Run: Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
#>

# =============================================================================
# CONFIGURATION
# =============================================================================

# Path to EML files (folder structure mirrors PST structure)
$SourcePath = "C:\Users\Public\Thunderbird_Staging"

# Isolated sandbox for temp file operations (prevents path errors)
$SandboxDir = "C:\Users\Public\OutlookImportTemp"
if (-not (Test-Path $SandboxDir)) { New-Item -ItemType Directory -Path $SandboxDir | Out-Null }
$TempStagingFile = Join-Path $SandboxDir "temp_import.eml"

# Outlook PST root folder name; must match name in Outlook sidebar exactly
$TargetPSTName = "Thunderbird_Archive"

# =============================================================================
# INITIALIZE OUTLOOK
# =============================================================================

$Outlook = New-Object -ComObject Outlook.Application
$NS = $Outlook.GetNamespace("MAPI")

try {
    $TargetRoot = $NS.Folders.Item($TargetPSTName)
} catch {
    Write-Error "Could not find a root folder named '$TargetPSTName' in Outlook sidebar."
    Write-Error "Create a local PST file in Outlook first, then name its root folder '$TargetPSTName'."
    Exit
}

# =============================================================================
# IMPORT FUNCTION
# =============================================================================

function Import-EMLFolder ($CurrentFSFolder, $CurrentOutlookFolder) {
    if ($null -eq $CurrentOutlookFolder) { return }

    Write-Host "`n[FOLDER] Processing: $($CurrentFSFolder.Name)" -ForegroundColor Cyan

    # --- Parse and inject emails ---
    $Files = Get-ChildItem -Path $CurrentFSFolder.FullName -Filter "*.eml"
    $FileCount = 0

    foreach ($File in $Files) {
        try {
            # Copy to sandbox to avoid path length issues
            Copy-Item $File.FullName -Destination $TempStagingFile -Force -ErrorAction Stop

            # Load EML bytes into ADODB stream
            $AdoStream = New-Object -ComObject "ADODB.Stream"
            $AdoStream.Type = 1  # Binary
            $AdoStream.Open()
            $AdoStream.LoadFromFile($TempStagingFile)

            # Parse EML via CDO
            $CdoMsg = New-Object -ComObject "CDO.Message"
            $CdoMsg.DataSource.OpenObject($AdoStream, "_Stream")

            # Create PostItem (Type 6), which works in any PST folder
            $MailItem = $CurrentOutlookFolder.Items.Add(6)

            # Force it to behave as standard email
            $MailItem.MessageClass = "IPM.Note"

            # Subject: fall back to filename if empty
            $DisplaySubject = if (![string]::IsNullOrEmpty($CdoMsg.Subject)) { $CdoMsg.Subject } else { $File.BaseName }
            $MailItem.Subject = $DisplaySubject

            # Body: prefer HTML, fall back to plain text
            if (![string]::IsNullOrEmpty($CdoMsg.HTMLBody)) {
                $MailItem.HTMLBody = $CdoMsg.HTMLBody
            } else {
                $MailItem.Body = $CdoMsg.TextBody
            }

            # --- Map date ---
            $DateVal = $CdoMsg.Fields.Item("urn:schemas:mailheader:date").Value
            if ($null -ne $DateVal) {
                try { $MailItem.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0E060040", $DateVal) } catch {}
            }

            # --- Map sender ---
            $From = $CdoMsg.From
            if (![string]::IsNullOrEmpty($From)) {
                try {
                    $MailItem.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0C1A001E", $From)
                    $MailItem.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0C1F001E", $From)
                } catch {}
            }

            # --- Map attachments ---
            if ($CdoMsg.Attachments.Count -gt 0) {
                foreach ($CdoAtt in $CdoMsg.Attachments) {
                    $AttName = $CdoAtt.FileName
                    if ([string]::IsNullOrWhiteSpace($AttName)) { $AttName = "asset_" + [guid]::NewGuid().ToString().Substring(0,8) }
                    $CleanAttName = $AttName -replace '[\\/:*?"<>|]', '_'
                    $TempAttPath = Join-Path $SandboxDir $CleanAttName
                    try {
                        $CdoAtt.SaveToFile($TempAttPath)
                        $MailItem.Attachments.Add($TempAttPath) | Out-Null
                    } catch {} finally {
                        if (Test-Path $TempAttPath) { Remove-Item $TempAttPath -Force }
                    }
                }
            }

            # Mark as read (historical item)
            try {
                $MailItem.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0E070003", 1)
            } catch {}

            # Save to PST local storage
            $MailItem.Save()

            $FileCount++
            Write-Host "  [OK] $DisplaySubject" -ForegroundColor Green

            # Release COM objects to prevent memory leaks
            $AdoStream.Close()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($AdoStream) | Out-Null
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($CdoMsg) | Out-Null
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($MailItem) | Out-Null

        } catch {
            Write-Warning "  [FAIL] Skipping: $($File.Name). Error: $_"
        } finally {
            if (Test-Path $TempStagingFile) { Remove-Item $TempStagingFile -Force }
        }
    }

    Write-Host "Loaded $FileCount emails to PST folder: $($CurrentOutlookFolder.Name)" -ForegroundColor Gray

    # --- Replicate subfolders recursively ---
    $SubDirs = Get-ChildItem -Path $CurrentFSFolder.FullName -Directory
    foreach ($SubDir in $SubDirs) {
        $CleanName = $SubDir.Name -replace "\.sbd$", ""
        $CleanName = $CleanName -replace '[\\/:*?"<>|\[\]]', '_'
        $CleanName = $CleanName.Trim()
        if ([string]::IsNullOrWhiteSpace($CleanName)) { $CleanName = "Unresolved_Folder" }

        $NextOutlookFolder = $null
        try {
            $NextOutlookFolder = $CurrentOutlookFolder.Folders.Item($CleanName)
        } catch {
            try { $NextOutlookFolder = $CurrentOutlookFolder.Folders.Add($CleanName) } catch {}
        }
        Import-EMLFolder $SubDir $NextOutlookFolder
    }
}

# =============================================================================
# RUN
# =============================================================================

Write-Host "Starting Thunderbird to Outlook import..." -ForegroundColor Yellow
Write-Host "Source: $SourcePath" -ForegroundColor Gray
Write-Host "Target: $TargetPSTName" -ForegroundColor Gray
Write-Host "Make sure Classic Outlook is open!" -ForegroundColor Yellow

$RootFSFolder = Get-Item $SourcePath
Import-EMLFolder $RootFSFolder $TargetRoot

# Cleanup sandbox
if (Test-Path $SandboxDir) { Remove-Item $SandboxDir -Recurse -Force -ErrorAction SilentlyContinue }

Write-Host "`nMigration complete! Check your $TargetPSTName folders in Outlook." -ForegroundColor Green

param(
    [Parameter(Mandatory = $true)][string]$SourcePath,
    [Parameter(Mandatory = $true)][string]$TargetPSTName,
    [Parameter(Mandatory = $true)][string]$CancelFile,
    [switch]$MarkAsRead
)

$ErrorActionPreference = "Stop"
$SandboxDir = Join-Path ([IO.Path]::GetTempPath()) ("mbox2pst-" + [guid]::NewGuid())
New-Item -ItemType Directory -Path $SandboxDir -Force | Out-Null
$TempStagingFile = Join-Path $SandboxDir "message.eml"
$Imported = 0
$Failed = 0
$Current = 0
$Cancelled = $false

function Release-ComObject($Object) {
    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

function Clean-OutlookFolderName([string]$Name) {
    $CleanName = $Name -replace '\.sbd$', ''
    $CleanName = $CleanName -replace '[\\/:*?"<>|\[\]]', '_'
    $CleanName = $CleanName.Trim()
    if ([string]::IsNullOrWhiteSpace($CleanName)) { return "Unresolved_Folder" }
    return $CleanName
}

function Get-OrCreateOutlookFolder($Parent, [string]$Name) {
    $Folder = $null
    try { $Folder = $Parent.Folders.Item($Name) } catch {}
    if ($null -eq $Folder) { $Folder = $Parent.Folders.Add($Name) }
    return $Folder
}

function Import-EMLFolder([IO.DirectoryInfo]$CurrentFSFolder, $CurrentOutlookFolder, [bool]$IsRoot = $false) {
    if (Test-Path $CancelFile) {
        $script:Cancelled = $true
        return
    }

    $TargetFolder = $CurrentOutlookFolder
    $ReleaseTargetFolder = $false
    if ($IsRoot) {
        $RootFiles = @(Get-ChildItem -LiteralPath $CurrentFSFolder.FullName -Filter "*.eml" -File)
        if ($RootFiles.Count -gt 0) {
            $TargetFolder = Get-OrCreateOutlookFolder $CurrentOutlookFolder (Clean-OutlookFolderName $CurrentFSFolder.Name)
            $ReleaseTargetFolder = $true
        }
    }

    $Files = @(Get-ChildItem -LiteralPath $CurrentFSFolder.FullName -Filter "*.eml" -File | Sort-Object Name)
    foreach ($File in $Files) {
        if (Test-Path $CancelFile) {
            $script:Cancelled = $true
            return
        }
        $script:Current++
        $AdoStream = $null
        $CdoMsg = $null
        $MailItem = $null
        try {
            Copy-Item -LiteralPath $File.FullName -Destination $TempStagingFile -Force
            $AdoStream = New-Object -ComObject "ADODB.Stream"
            $AdoStream.Type = 1
            $AdoStream.Open()
            $AdoStream.LoadFromFile($TempStagingFile)

            $CdoMsg = New-Object -ComObject "CDO.Message"
            $CdoMsg.DataSource.OpenObject($AdoStream, "_Stream")
            $MailItem = $TargetFolder.Items.Add(6)
            $MailItem.MessageClass = "IPM.Note"
            $MailItem.Subject = if (![string]::IsNullOrWhiteSpace($CdoMsg.Subject)) { $CdoMsg.Subject } else { $File.BaseName }

            if (![string]::IsNullOrWhiteSpace($CdoMsg.HTMLBody)) { $MailItem.HTMLBody = $CdoMsg.HTMLBody }
            else { $MailItem.Body = $CdoMsg.TextBody }

            try {
                $DateVal = $CdoMsg.Fields.Item("urn:schemas:mailheader:date").Value
                if ($null -ne $DateVal) {
                    $MailItem.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0E060040", $DateVal)
                }
            } catch {}

            try {
                $From = $CdoMsg.From
                if (![string]::IsNullOrWhiteSpace($From)) {
                    $MailItem.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0C1A001E", $From)
                    $MailItem.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0C1F001E", $From)
                }
            } catch {}

            foreach ($CdoAtt in @($CdoMsg.Attachments)) {
                $TempAttPath = $null
                try {
                    $AttName = if ([string]::IsNullOrWhiteSpace($CdoAtt.FileName)) { "asset_" + [guid]::NewGuid().ToString("N") } else { $CdoAtt.FileName }
                    $AttName = $AttName -replace '[\\/:*?"<>|]', '_'
                    $TempAttPath = Join-Path $SandboxDir ([guid]::NewGuid().ToString("N") + "_" + $AttName)
                    $CdoAtt.SaveToFile($TempAttPath)
                    [void]$MailItem.Attachments.Add($TempAttPath)
                } finally {
                    if ($TempAttPath -and (Test-Path $TempAttPath)) { Remove-Item -LiteralPath $TempAttPath -Force }
                    Release-ComObject $CdoAtt
                }
            }

            if ($MarkAsRead) {
                try { $MailItem.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0E070003", 1) } catch {}
            }
            $MailItem.Save()
            $script:Imported++
            Write-Output ("PROGRESS|{0}|{1}|{2}|{3}" -f $script:Current, $script:Imported, $script:Failed, $CurrentFSFolder.Name)
        } catch {
            $script:Failed++
            Write-Output ("FAILED|{0}" -f ($_.Exception.Message -replace '[\r\n|]', ' '))
            Write-Output ("PROGRESS|{0}|{1}|{2}|{3}" -f $script:Current, $script:Imported, $script:Failed, $CurrentFSFolder.Name)
        } finally {
            if ($null -ne $AdoStream) { try { $AdoStream.Close() } catch {} }
            Release-ComObject $MailItem
            Release-ComObject $CdoMsg
            Release-ComObject $AdoStream
            if (Test-Path $TempStagingFile) { Remove-Item -LiteralPath $TempStagingFile -Force }
        }
    }

    if ($script:Cancelled) {
        if ($ReleaseTargetFolder) { Release-ComObject $TargetFolder }
        return
    }

    foreach ($SubDir in @(Get-ChildItem -LiteralPath $CurrentFSFolder.FullName -Directory | Sort-Object Name)) {
        if (Test-Path $CancelFile) {
            $script:Cancelled = $true
            break
        }
        $NextFolder = $null
        try {
            $NextFolder = Get-OrCreateOutlookFolder $TargetFolder (Clean-OutlookFolderName $SubDir.Name)
            Import-EMLFolder $SubDir $NextFolder
            if ($script:Cancelled) { break }
        } finally {
            Release-ComObject $NextFolder
        }
    }
    if ($ReleaseTargetFolder) { Release-ComObject $TargetFolder }
}

$Outlook = $null
$Namespace = $null
$TargetRoot = $null
try {
    Write-Output "STATUS|Connecting to Classic Outlook…"
    $Outlook = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
    $Namespace = $Outlook.GetNamespace("MAPI")
    try { $TargetRoot = $Namespace.Folders.Item($TargetPSTName) } catch {
        throw "Outlook does not have an open PST named '$TargetPSTName'."
    }
    $RootFolder = Get-Item -LiteralPath $SourcePath
    Import-EMLFolder $RootFolder $TargetRoot $true
    if ($script:Cancelled) { Write-Output "CANCELLED" }
    else { Write-Output ("COMPLETE|{0}|{1}" -f $script:Imported, $script:Failed) }
} finally {
    Release-ComObject $TargetRoot
    Release-ComObject $Namespace
    Release-ComObject $Outlook
    if (Test-Path $SandboxDir) { Remove-Item -LiteralPath $SandboxDir -Recurse -Force -ErrorAction SilentlyContinue }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

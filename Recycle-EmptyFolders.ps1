# --- Script: Recycle-EmptyFolders.ps1 ---

<#
.SYNOPSIS
    Finds and optionally moves empty folders within a specified target directory to the Recycle Bin.

.DESCRIPTION
    This script scans a given directory for subdirectories that are completely empty (containing no files or folders).
    It can either list these folders (Dry Run) or move them to the Windows Recycle Bin.

.PARAMETER TargetFolderPath
    The full path to the folder that will be scanned for empty subdirectories.
    This script looks *inside* this folder for empty folders, not the folder itself.

.PARAMETER DryRun
    If specified, the script will only list the empty folders it finds without moving them.
    Use this to preview which folders would be affected.

.EXAMPLE
    # List empty folders in the target directory (Dry Run)
    .\Recycle-EmptyFolders.ps1 -TargetFolderPath "C:\Path\To\Target" -DryRun

.EXAMPLE
    # Move empty folders in the target directory to the Recycle Bin
    .\Recycle-EmptyFolders.ps1 -TargetFolderPath "C:\Path\To\Target"

.NOTES
    Requires .NET Framework/Library access to Microsoft.VisualBasic for Recycle Bin operations.
    Uses [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory for stability.
#>
param(
    [Parameter(Mandatory = $true, HelpMessage = "Enter the full path to the target folder to scan for empty subfolders.")]
    [string]$TargetFolderPath,
    [Parameter(Mandatory = $false)]
    [switch]$DryRun
)

# --- 1. Validate Input Path ---
Write-Host "Checking target folder path: $TargetFolderPath"
if (-not (Test-Path $TargetFolderPath -PathType Container)) {
    Write-Error "Target folder path does not exist or is not a directory: $TargetFolderPath"
    exit 1
}
# Ensure path ends with \ for Get-ChildItem consistency
if (-not $TargetFolderPath.EndsWith('\')) {
    $TargetFolderPath += '\'
}
Write-Host "Target folder path validated."

# --- 2. Load Required .NET Assembly ---
Write-Host "Loading Microsoft.VisualBasic assembly for Recycle Bin operations..."
try {
    Add-Type -AssemblyName "Microsoft.VisualBasic" -ErrorAction Stop
    Write-Host "Microsoft.VisualBasic assembly loaded successfully."
} catch {
    Write-Error "Failed to load 'Microsoft.VisualBasic' assembly. Error: $_"
    Write-Host "This script requires this assembly to move folders to the Recycle Bin."
    exit 1
}

# --- 3. Find Empty Folders ---
Write-Host "`nScanning '$TargetFolderPath' for empty subfolders..."
try {
    # Get all directories recursively within the target folder
    $AllSubFolders = Get-ChildItem -Path $TargetFolderPath -Directory -Recurse -ErrorAction Stop | Sort-Object FullName -Descending

    # Filter for truly empty ones (no files or subdirectories inside)
    $EmptyFolders = $AllSubFolders | Where-Object {
        $items = Get-ChildItem -Path $_.FullName -Force -ErrorAction SilentlyContinue
        return ($null -eq $items -or $items.Count -eq 0)
    }
} catch {
    Write-Error "An error occurred while scanning for folders: $_"
    exit 1
}

$EmptyFolderCount = $EmptyFolders.Count
Write-Host "Scan complete. Found $EmptyFolderCount empty subfolder(s)."

# --- 4. Process Empty Folders ---
if ($EmptyFolderCount -eq 0) {
    Write-Host "No empty folders found in '$TargetFolderPath'. Nothing to do."
    exit 0
}

if ($DryRun) {
    Write-Host "`n[DRY-RUN] The following empty folders WOULD BE moved to the Recycle Bin:" -ForegroundColor Cyan
    $EmptyFolders | ForEach-Object { Write-Host "  $($_.FullName)" }
    Write-Host "`n[DRY-RUN] No file system changes were made." -ForegroundColor Cyan
} else {
    Write-Host "`nSending $EmptyFolderCount identified empty folder(s) to the Recycle Bin..."
    $MovedCount = 0
    $ErrorCount = 0

    foreach ($Folder in $EmptyFolders) {
        $FolderPath = $Folder.FullName
        Write-Host "Attempting to move: $FolderPath"
        try {
            # Use .NET VB method to move directory to Recycle Bin
            # UIOption: OnlyErrorDialogs means no confirmation, just error dialogs if it fails.
            # RecycleOption: SendToRecycleBin
            # UICancelOption: DoNothing means if UI pops up (it shouldn't), cancel is ignored.
            [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory($FolderPath, 'OnlyErrorDialogs', 'SendToRecycleBin', 'DoNothing')
            Write-Host "  -> Success" -ForegroundColor Green
            $MovedCount++
        } catch [System.UnauthorizedAccessException] {
            Write-Error "  -> Access denied moving '$FolderPath'. Error: $($_.Exception.Message)" -ErrorAction Continue
            $ErrorCount++
        } catch [System.IO.IOException] {
            # This includes directory-in-use errors
            Write-Error "  -> IO Error moving '$FolderPath'. It might be in use. Error: $($_.Exception.Message)" -ErrorAction Continue
            $ErrorCount++
        } catch [System.Management.Automation.MethodInvocationException] {
            # Catch errors specifically from the .NET VB method call
            Write-Error "  -> Error calling VB DeleteDirectory for '$FolderPath'. Error: $($_.Exception.InnerException.Message)" -ErrorAction Continue
            $ErrorCount++
        } catch {
            Write-Error "  -> Unexpected error moving '$FolderPath'. Error: $_" -ErrorAction Continue
            $ErrorCount++
        }
    }

    # --- 5. Report Results ---
    Write-Host "`n--- Empty Folder Cleanup Summary ---"
    Write-Host "Folders moved to Recycle Bin: $MovedCount"
    if ($ErrorCount -gt 0) {
        Write-Warning "Errors encountered: $ErrorCount"
    }
    Write-Host "------------------------------`n"
}

Write-Host "Script execution finished."

# --- Script: Deduplicate-TargetFolder.ps1 ---

param(
    [Parameter(Mandatory = $true)]
    [string]$ReferenceFolderPath,
    [Parameter(Mandatory = $true)]
    [string]$TargetFolderPath,
    [Parameter(Mandatory = $false)] # Optional: Specify output file name
    [string]$OutputJsonFile = "duplicates_output.json",
    [Parameter(Mandatory = $false)]
    [switch]$DryRun # New switch for dry-run mode
)

# --- 1. Validate Input Paths ---
if (-not (Test-Path $ReferenceFolderPath -PathType Container)) {
    Write-Error "Reference folder path does not exist or is not a directory: $ReferenceFolderPath"
    exit 1
}
if (-not (Test-Path $TargetFolderPath -PathType Container)) {
    Write-Error "Target folder path does not exist or is not a directory: $TargetFolderPath"
    exit 1
}

# Ensure paths end with a backslash for consistent comparison later
if (-not $ReferenceFolderPath.EndsWith('\')) { $ReferenceFolderPath += '\' }
if (-not $TargetFolderPath.EndsWith('\')) { $TargetFolderPath += '\' }

# --- 2. Define Path to czkawka-CLI.exe ---
# Update this path to where your executable is located
$CzkawkaCLIPath = ".\czkawka-CLI.exe" # Assuming it's in the current directory

# --- Helper Function to Run czkawka and Check for Intra-Folder Duplicates ---
function Check-IntraFolderDuplicates {
    param(
        [string]$FolderPath,
        [string]$Description,
        [bool]$IsDryRun # Accept DryRun status as a parameter
    )

    Write-Host "Checking for duplicates within $Description folder: $FolderPath"
    $TempOutputFile = [System.IO.Path]::GetRandomFileName() + ".json"

    # Prepare arguments for czkawka. PowerShell handles quoting paths passed as array elements.
    $CheckArgs = @(
        "dup",
        "--search-method", "HASH",
        "--hash-type", "BLAKE3"
    )
    # Add czkawka's --dry-run flag if our script is in dry-run mode
    if ($IsDryRun) {
        $CheckArgs += "--dry-run"
    }
    $CheckArgs += @(
        "--directories", $FolderPath,
        "--compact-file-to-save", $TempOutputFile
    )

    try {
        Write-Host "Executing: & $CzkawkaCLIPath $($CheckArgs -join ' ')"

        # Capture output from czkawka
        $CzkawkaOutput = & $CzkawkaCLIPath $CheckArgs 2>&1
        # Display the output as it happens
        $CzkawkaOutput | ForEach-Object { Write-Host $_ }

        # Logic to detect duplicates based on mode
        if ($IsDryRun) {
            # In dry-run mode, check the captured console output for the "Found X duplicated files" message AS A FALLBACK
            # BUT, prefer to check the JSON file if it was created.
            $DuplicationFoundInConsole = $CzkawkaOutput -match "Found \d+ duplicated files"

            # Check if the JSON file was actually created (it should be now)
            if (Test-Path $TempOutputFile -PathType Leaf) {
                Write-Host "[DRY-RUN] Temporary JSON output file '$TempOutputFile' was created."
                $TempData = Get-Content -Path $TempOutputFile -Raw | ConvertFrom-Json
                Remove-Item -Path $TempOutputFile -Force -ErrorAction SilentlyContinue # Clean up temp file for dry-run

                # Check JSON data for duplicates
                if (($TempData -ne $null) -and ($TempData.Count -gt 0)) {
                    $LooksLikeGroup = $TempData[0].PSObject.Properties.Name -contains 'files' -or $TempData[0].PSObject.Properties.Name -contains 'size'
                    if ($LooksLikeGroup) {
                        # Cleaner error output for dry-run (based on JSON)
                        Write-Host "ERROR: Duplicates found within the $Description folder." -ForegroundColor Red
                        Write-Host "       Path: $FolderPath" -ForegroundColor Red
                        Write-Host "       Please resolve these duplicates before running this script." -ForegroundColor Red
                        Write-Host "       You can use 'czkawka-CLI.exe dup --directories `"$FolderPath`"' to find them." -ForegroundColor Yellow
                        exit 1 # Exit the script with error code 1
                    } else {
                         Write-Warning "[DRY-RUN] JSON output exists but first item doesn't look like a group. Assuming no duplicates. (Properties: $($TempData[0].PSObject.Properties.Name -join ', '))"
                    }
                } else {
                    Write-Host "[DRY-RUN] No duplicates found within $Description folder '$FolderPath' (JSON output empty/null)."
                }
            } else {
                # Fallback: If no JSON file, rely on console parsing (less reliable)
                Write-Warning "[DRY-RUN] Temporary JSON output file '$TempOutputFile' was NOT created. Falling back to console parsing."
                if ($DuplicationFoundInConsole) {
                    Write-Host "ERROR: Duplicates found within the $Description folder (detected in console output)." -ForegroundColor Red
                    Write-Host "       Path: $FolderPath" -ForegroundColor Red
                    Write-Host "       Please resolve these duplicates before running this script." -ForegroundColor Red
                    Write-Host "       You can use 'czkawka-CLI.exe dup --directories `"$FolderPath`"' to find them." -ForegroundColor Yellow
                    exit 1
                } else {
                    Write-Host "[DRY-RUN] No duplicates detected in console output for $Description folder '$FolderPath' (fallback)."
                }
            }

        } else {
            # In normal mode, check if the JSON output file was created and analyze it (unchanged)
            if (Test-Path $TempOutputFile -PathType Leaf) {
                $TempData = Get-Content -Path $TempOutputFile -Raw | ConvertFrom-Json
                Remove-Item -Path $TempOutputFile -Force -ErrorAction SilentlyContinue # Clean up temp file
                # Check if any duplicates were found in the JSON data
                if (($TempData -ne $null) -and ($TempData.Count -gt 0)) {
                    $LooksLikeGroup = $TempData[0].PSObject.Properties.Name -contains 'files' -or $TempData[0].PSObject.Properties.Name -contains 'size'
                    if ($LooksLikeGroup) {
                        # Cleaner error output for normal mode
                        Write-Host "ERROR: Duplicates found within the $Description folder." -ForegroundColor Red
                        Write-Host "       Path: $FolderPath" -ForegroundColor Red
                        Write-Host "       Please resolve these duplicates before running this script." -ForegroundColor Red
                        Write-Host "       You can use 'czkawka-CLI.exe dup --directories `"$FolderPath`"' to find them." -ForegroundColor Yellow
                        exit 1 # Exit the script with error code 1
                    } else {
                         Write-Warning "JSON output exists but first item doesn't look like a duplicate group (properties: $($TempData[0].PSObject.Properties.Name -join ', ')). Assuming no duplicates."
                    }
                } else {
                    Write-Host "No duplicates found within $Description folder '$FolderPath' (JSON output empty/null)."
                }
            } else {
                Write-Warning "Check for $Description duplicates completed, but temporary output file '$TempOutputFile' was not found. Assuming no duplicates or a potential issue with czkawka."
            }
        }
    } catch {
        Write-Error "Failed to run intra-folder duplicate check for '$FolderPath'. Error: $_"
        # Clean up temp file if it exists
        if (Test-Path $TempOutputFile -PathType Leaf) {
            try { Remove-Item -Path $TempOutputFile -Force -ErrorAction SilentlyContinue } catch { Write-Warning "Could not remove temporary file '$TempOutputFile'" }
        }
        exit 1
    }
    Write-Host "No duplicates found within $Description folder (check passed). Proceeding..."
}


# --- 3. Check for Duplicates within Target Folder FIRST ---
Check-IntraFolderDuplicates -FolderPath $TargetFolderPath -Description "Target" -IsDryRun $DryRun.IsPresent

# --- 4. Check for Duplicates within Reference Folder (Optional but good practice) ---
# Check-IntraFolderDuplicates -FolderPath $ReferenceFolderPath -Description "Reference" -IsDryRun $DryRun.IsPresent


# --- 5. Run czkawka to find duplicates BETWEEN folders and save to JSON ---
# Prepare arguments for the main czkawka run
$CzkawkaArguments = @(
    "dup",
    "--search-method", "HASH",
    "--hash-type", "BLAKE3"
)
# Add czkawka's --dry-run flag if our script is in dry-run mode
if ($DryRun) {
    $CzkawkaArguments += "--dry-run"
}
$CzkawkaArguments += @(
    "--directories", $ReferenceFolderPath,
    "--directories", $TargetFolderPath,
    "--compact-file-to-save", $OutputJsonFile
)

Write-Host "`nRunning czkawka to find duplicates BETWEEN folders..."
Write-Host "Executing: & $CzkawkaCLIPath $($CzkawkaArguments -join ' ')"

# Execute czkawka (capture output for display)
$CzkawkaOutput = & $CzkawkaCLIPath $CzkawkaArguments 2>&1
$CzkawkaOutput | ForEach-Object { Write-Host $_ } # Show czkawka output

# --- 6. Check if output file was created ---
# This check is now valid for both normal and dry-run modes
if (-not (Test-Path $OutputJsonFile -PathType Leaf)) {
    Write-Host "czkawka completed. Output file '$OutputJsonFile' was not found."
    Write-Host "This usually means no duplicates were found, or czkawka encountered an issue."
    Write-Host "Check czkawka logs/messages above."
    # Set $DuplicatesDataRaw to $null to indicate no data to process
    $DuplicatesDataRaw = $null
    # Decide if you want to exit or continue assuming no dups.
    # Let's continue for now.
} else {
    Write-Host "czkawka output file '$OutputJsonFile' was created."
}

# --- 7. Read and Parse the JSON Output ---
# This step now works for both normal and dry-run modes if the file exists
if (Test-Path $OutputJsonFile -PathType Leaf) {
    Write-Host "Reading results from $OutputJsonFile..."
    try {
        # --- CHANGE: Load raw data ---
        $DuplicatesDataRaw = Get-Content -Path $OutputJsonFile -Raw | ConvertFrom-Json
        # --- END CHANGE ---
    } catch {
        Write-Error "Failed to read or parse JSON file '$OutputJsonFile'. Error: $_"
        # If parsing fails, we can't proceed safely.
        exit 1
    }
} else {
    # File doesn't exist
    Write-Host "Output file '$OutputJsonFile' confirmed absent. Assuming no duplicates."
    $DuplicatesDataRaw = $null
}

# --- 8. Identify Target Folder Duplicates (for reporting in DryRun, moving in Normal) ---
# This logic now correctly parses the dictionary structure
$TargetFilesToProcess = @() # Explicitly initialize as array

# --- CHANGE: Improved path matching logic ---
# Ensure TargetFolderPath ends with \ for consistent matching
$NormalizedTargetPath = $TargetFolderPath
if (-not $NormalizedTargetPath.EndsWith('\')) {
    $NormalizedTargetPath += '\'
}
# Convert to lowercase for case-insensitive comparison
$NormalizedTargetPathLower = $NormalizedTargetPath.ToLower()

if ($null -ne $DuplicatesDataRaw) {
    # $DuplicatesDataRaw is a PSCustomObject representing the dictionary
    # Iterate through its properties (each property is a Key-Value pair: Size -> GroupsArray)
    foreach ($SizeProperty in $DuplicatesDataRaw.PSObject.Properties) {
        # $SizeProperty.Name is the Size (e.g., "825973")
        # $SizeProperty.Value is the Array of Groups for that size
        $GroupsArrayForThisSize = $SizeProperty.Value

        # Iterate through each Group Array within the Groups Array for this size
        foreach ($GroupArray in $GroupsArrayForThisSize) {
            # $GroupArray is an array of file objects that are duplicates

            # --- CRITICAL FIX: Use normalized, lowercase paths for robust matching ---
            # Wrap the result in @() to force it to be an array, fixing potential PowerShell single-element array issues
            $TargetFilesInThisGroup = @($GroupArray | Where-Object {
                $_.path.ToLower() -like "$NormalizedTargetPathLower*"
            })
            # --- END CRITICAL FIX ---

            # If target files were found in this group, add them to the main list
            if ($TargetFilesInThisGroup.Count -gt 0) {
                $TargetFilesToProcess += $TargetFilesInThisGroup
            }
            # If no target files, this group is not relevant for moving/deleting from target
        }
    }
    Write-Host "Found $($TargetFilesToProcess.Count) duplicate file(s) in the TARGET folder based on JSON analysis."
} else {
     Write-Host "No cross-folder duplicates data found in JSON (it was null)."
}
# --- END CHANGE ---


# --- Perform Actions based on Mode ---
if ($DryRun) {
    # --- DRY-RUN MODE ---
    Write-Host "`n[DRY-RUN] Simulation completed." -ForegroundColor Cyan
    if ($TargetFilesToProcess.Count -gt 0) {
        Write-Host "[DRY-RUN] The following files in the TARGET folder WOULD BE moved to the Recycle Bin:" -ForegroundColor Cyan
        $TargetFilesToProcess | ForEach-Object { Write-Host "  $($_.path)" }
    } else {
        Write-Host "[DRY-RUN] No duplicate files found in the TARGET folder." -ForegroundColor Green
    }
    Write-Host "[DRY-RUN] No file system changes were made." -ForegroundColor Cyan
    Write-Host "[DRY-RUN] Output file '$OutputJsonFile' was created and will be left for inspection." -ForegroundColor Yellow
    # Optional: Remove the output file at the end of dry-run if desired
    # Remove-Item -Path $OutputJsonFile -Force -ErrorAction SilentlyContinue
    Write-Host "`n------------------------------`n" -ForegroundColor Cyan

} else {
    # --- NORMAL MODE ---
    if ($TargetFilesToProcess.Count -gt 0) {
        Write-Host "`nMoving identified target folder duplicate files to Recycle Bin..."
        $MovedFilesCount = 0
        $ErrorCount = 0
        $TotalFiles = $TargetFilesToProcess.Count
        $BatchSize = 100 # Define the batch size (can keep this for progress reporting even if not for COM delay)
        $UseVBMethod = $true # Flag to indicate which method is being used

        # --- CHANGE 1: Load VisualBasic Assembly ---
        try {
            Add-Type -AssemblyName "Microsoft.VisualBasic" -ErrorAction Stop
            Write-Host "Using .NET Microsoft.VisualBasic method for Recycle Bin operation." -ForegroundColor Green
        } catch {
            Write-Warning "Failed to load 'Microsoft.VisualBasic' assembly. Error: $_"
            Write-Host "Falling back to permanent file deletion using [System.IO.File]::Delete()" -ForegroundColor Yellow
            try {
                Add-Type -AssemblyName "System.IO" -ErrorAction Stop # Usually core, but good to ensure
            } catch {
                Write-Error "Failed to ensure 'System.IO' is available. Cannot proceed. Error: $_"
                exit 1
            }
            $UseVBMethod = $false
        }
        # --- END CHANGE 1 ---

        # --- CHANGE 2: Process files (in batches for progress, but not for COM delay) ---
        for ($i = 0; $i -lt $TotalFiles; $i += $BatchSize) {
            # Calculate the end index for this batch (mainly for progress display)
            $endIndex = [Math]::Min(($i + $BatchSize - 1), ($TotalFiles - 1))
            # Extract the current batch (for progress display)
            $CurrentBatch = $TargetFilesToProcess[$i..$endIndex]
            $BatchNumber = [Math]::Floor($i / $BatchSize) + 1
            $FilesInBatch = $CurrentBatch.Count

            Write-Host "`n--- Processing Batch $BatchNumber (Files $($i+1) to $($endIndex+1) of $TotalFiles) ---"

            # Process each file in the current "batch"
            $BatchIndex = 0
            foreach ($FileToRemove in $CurrentBatch) {
                $BatchIndex++
                $GlobalIndex = $i + $BatchIndex
                $FilePath = $FileToRemove.path
                Write-Host "[$GlobalIndex/$TotalFiles] [Batch $BatchNumber - Item $BatchIndex/$FilesInBatch] Attempting to move: $FilePath"

                try {
                    if ($UseVBMethod) {
                        # --- CHANGE 3: Use .NET VisualBasic Method ---
                        # UIOption: OnlyErrorDialogs means no confirmation, just error dialogs if it fails.
                        # RecycleOption: SendToRecycleBin
                        # UICancelOption: DoNothing means if UI pops up (it shouldn't), cancel is ignored.
                        [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($FilePath, 'OnlyErrorDialogs', 'SendToRecycleBin', 'DoNothing')
                        # --- END CHANGE 3 ---
                    } else {
                        # Fallback: Permanent deletion using .NET
                        [System.IO.File]::Delete($FilePath)
                    }
                    Write-Host "  -> Success"
                    $MovedFilesCount++

                    # Optional: Very small delay between files if needed even with .NET method
                    # Start-Sleep -Milliseconds 10

                } catch [System.UnauthorizedAccessException] {
                    Write-Error "[$GlobalIndex/$TotalFiles] [Batch $BatchNumber - Item $BatchIndex/$FilesInBatch] Access denied moving '$FilePath'. Error: $($_.Exception.Message)" -ErrorAction Continue
                    $ErrorCount++
                } catch [System.IO.IOException] {
                    # This includes file-in-use errors
                    Write-Error "[$GlobalIndex/$TotalFiles] [Batch $BatchNumber - Item $BatchIndex/$FilesInBatch] IO Error moving '$FilePath'. It might be in use. Error: $($_.Exception.Message)" -ErrorAction Continue
                    $ErrorCount++
                } catch [System.Management.Automation.MethodInvocationException] {
                    # Catch errors specifically from the .NET VB method call
                    Write-Error "[$GlobalIndex/$TotalFiles] [Batch $BatchNumber - Item $BatchIndex/$FilesInBatch] Error calling VB DeleteFile for '$FilePath'. Error: $($_.Exception.InnerException.Message)" -ErrorAction Continue
                    $ErrorCount++
                } catch {
                    Write-Error "[$GlobalIndex/$TotalFiles] [Batch $BatchNumber - Item $BatchIndex/$FilesInBatch] Unexpected error moving '$FilePath'. Error: $_" -ErrorAction Continue
                    $ErrorCount++
                }
            } # End of foreach file in current "batch"

            # Optional: Delay between "batches" if needed, even with .NET method
            # (This is now just a pause between groups of progress messages, not for COM stability)
            if (($i + $BatchSize) -lt $TotalFiles) {
                 Write-Host "[Batch $BatchNumber finished. Brief pause before next group...]"
                 Start-Sleep -Milliseconds 1000 # Adjust as needed
            }

        } # End of for loop iterating through "batches"
        # --- END CHANGE 2 ---


        # --- 9. Report Results (Normal Mode) ---
        $MethodUsed = if ($UseVBMethod) { "moved to Recycle Bin (via .NET VB)" } else { "permanently deleted (fallback)" }
        Write-Host "`n--- Deduplication Summary ---"
        Write-Host "Files ${MethodUsed}: $MovedFilesCount" # --- CORRECTED LINE ---
        if ($ErrorCount -gt 0) {
            Write-Warning "Errors encountered: $ErrorCount"
        }
        Write-Host "Original czkawka results saved to: $OutputJsonFile"
        Write-Host "Review remaining files in '$TargetFolderPath' and consider adding new unique ones to '$ReferenceFolderPath'."
        Write-Host "------------------------------`n"

    } else {
         Write-Host "No files to move in normal mode."
    }
    # Optional: Remove the temporary JSON file if not needed afterwards
    # Remove-Item -Path $OutputJsonFile -Force -ErrorAction SilentlyContinue
}

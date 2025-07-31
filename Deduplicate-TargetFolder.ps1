# --- Script: Deduplicate-TargetFolder.ps1 ---
# DEFECT: Crashed when it tries to process over 500 files to the Recycling Bin.

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
        $CurrentFileIndex = 0
        $TotalFiles = $TargetFilesToProcess.Count

        foreach ($FileToRemove in $TargetFilesToProcess) {
             $CurrentFileIndex++
             $FilePath = $FileToRemove.path
             Write-Host "[$CurrentFileIndex/$TotalFiles] Attempting to move: $FilePath"
             # Initialize COM objects as $null for this iteration
             $Shell = $null
             $Item = $null
             try {
                 # Move to Recycle Bin using Shell.Application COM object
                 $Shell = New-Object -ComObject Shell.Application
                 $Item = $Shell.Namespace(0).ParseName($FilePath)
                 if ($Item) {
                     $Item.InvokeVerb('delete') # This moves to Recycle Bin
                     Write-Host "  -> Success" # Indicate success
                     $MovedFilesCount++
                 } else {
                     Write-Warning "Could not find item in shell namespace to delete: $FilePath"
                     $ErrorCount++
                 }

                 # Explicitly release COM objects
                 if ($Item) {
                     [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Item) | Out-Null
                     $Item = $null # Set to null after releasing
                 }
                 if ($Shell) {
                     [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Shell) | Out-Null
                     $Shell = $null # Set to null after releasing
                 }

                 # Add a very brief delay to reduce stress on the Shell COM object
                 Start-Sleep -Milliseconds 5

             } catch [System.Runtime.InteropServices.COMException] {
                 # Catch specific COM exceptions (including potential ACCESS_VIOLATION)
                 Write-Error "[$CurrentFileIndex/$TotalFiles] COM Error moving file '$FilePath'. Error: $($_.Exception.Message) (HRESULT: $($_.Exception.HResult))" -ErrorAction Continue
                 $ErrorCount++
                 # Ensure COM objects are released even if an error occurred
                 if ($Item) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Item) | Out-Null }
                 if ($Shell) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Shell) | Out-Null }
             } catch [System.UnauthorizedAccessException] {
                 # Catch permission errors
                 Write-Error "[$CurrentFileIndex/$TotalFiles] Access denied moving file '$FilePath'. Error: $($_.Exception.Message)" -ErrorAction Continue
                 $ErrorCount++
                 # Ensure COM objects are released even if an error occurred
                 if ($Item) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Item) | Out-Null }
                 if ($Shell) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Shell) | Out-Null }
             } catch [System.IO.IOException] {
                 # Catch file system errors (e.g., file in use)
                 Write-Error "[$CurrentFileIndex/$TotalFiles] IO Error moving file '$FilePath'. Error: $($_.Exception.Message)" -ErrorAction Continue
                 $ErrorCount++
                 # Ensure COM objects are released even if an error occurred
                 if ($Item) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Item) | Out-Null }
                 if ($Shell) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Shell) | Out-Null }
             } catch {
                 # Catch any other unexpected errors
                 Write-Error "[$CurrentFileIndex/$TotalFiles] Unexpected error moving file '$FilePath'. Error: $_" -ErrorAction Continue
                 $ErrorCount++
                 # Ensure COM objects are released even if an unexpected error occurred
                 if ($Item) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Item) | Out-Null }
                 if ($Shell) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Shell) | Out-Null }
             } finally {
                 # Final safeguard to release COM objects if not already done in catch blocks
                 # Note: In normal flow, they should be released in the try block.
                 # This handles cases where an error bypassed the specific catches but objects were created.
                 if ($Item -ne $null) {
                     try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Item) | Out-Null } catch { Write-Debug "Failed to release Item COM object in Finally block." }
                 }
                 if ($Shell -ne $null) {
                     try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Shell) | Out-Null } catch { Write-Debug "Failed to release Shell COM object in Finally block." }
                 }
             }
        }
        # --- 9. Report Results (Normal Mode) ---
        Write-Host "`n--- Deduplication Summary ---"
        Write-Host "Files moved to Recycle Bin: $MovedFilesCount"
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

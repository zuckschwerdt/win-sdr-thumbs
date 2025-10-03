# SCRIPT REQUIRES: Microsoft.WindowsAPICodePack.Shell.dll
# Can be found here: https://github.com/ThioJoe/Lost-And-Found

# To allow running this script run this. (It only temporarily changes the policy for the current PowerShell session)
# Set-ExecutionPolicy -ExecutionPolicy unrestricted -Scope Process

# -------------------- DEFAULT SETTINGS --------------------
param(
    # Directory containing SVG files to test on
    [string]$dir,

    # Preset for thumbnail size, can be Small, Medium, Large, or ExtraLarge
    [ValidateSet("Small", "Medium", "Large", "ExtraLarge")]
    [string]$size = "Large", # Possible values: Small, Medium, Large, ExtraLarge

    # Number of times to run the benchmark and average the results
    [int]$numberOfRuns = 5,

    # Set your desired retrieval flags here
    [bool]$thumbnailOnly = $true,
    [bool]$inCacheOnly = $false,

    # Output a couple thumbnail result samples next to script to ensure thumbnails are actually rendering
    [bool]$outputSampleFiles = $false,

    # Path to the Windows API Code Pack "Microsoft.WindowsAPICodePack.Shell.dll", if not provided it defaults to looking in the script's directory
    # The DLL must be next to the other DLLS
    [Parameter(Mandatory = $false)]
    [string]$ApiPackDllPath = $null
)

# Set default for $dir if not provided, which is next to the script in a folder named "TestSvgFiles"
if (-not $dir -or $dir -eq "") {
    $dir = Join-Path -Path $PSScriptRoot -ChildPath "TestSvgFiles"
}


# ------------ Some Validation -----------------
if (-not (Test-Path $dir)) {
    Write-Error "Test directory '$dir' does not exist. Please provide a valid path containing SVG files to test on."
    return
}

if (-not $ApiPackDllPath) {
    # Default to the script's directory if no DLL path is provided
    $ApiPackDllPath = Join-Path -Path $PSScriptRoot -ChildPath "Windows API Code Pack 1.1\binaries\Microsoft.WindowsAPICodePack.Shell.dll"
    $userProvidedDllPathBool = $true
    # Check if the default DLL path exists
    if (-not (Test-Path $ApiPackDllPath)) {
        Write-Error @"
You must provide a valid path to the Windows API Code Pack DLL. Download it from the link at the top of the script and place the "Windows API Code Pack 1.1" folder next to the script, or provide a valid path to "Microsoft.WindowsAPICodePack.Shell.dll" using the -ApiPackDllPath parameter.
"@
        return
    }
} else {
    # If a path is provided, ensure it exists
    if (-not (Test-Path $ApiPackDllPath)) {
        Write-Error "Provided API Pack DLL path '$ApiPackDllPath' does not exist. Please check the path."
        return
    }
}



# -------------------- END SETTINGS --------------------

function Benchmark-Thumbnails {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [ValidateSet("Small", "Medium", "Large", "ExtraLarge")]
        [string]$ThumbnailPreset,

        [Parameter(Mandatory = $true)]
        [bool]$ThumbnailOnly,

        [Parameter(Mandatory = $true)]
        [bool]$InCacheOnly,

        [Parameter(Mandatory = $false)]
        [string]$Filter = '*.svg',

        [Parameter(Mandatory = $false)]
        [string]$ApiPackDllPath = $null
    )

    # --- Validate parameters ---

    Try {
        Add-Type -Path $ApiPackDllPath
    }
    Catch {
        Write-Error "Failed to load a required DLL. Ensure the file is not blocked."
        return
    }

    $fileList = Get-ChildItem -Path $Path -Filter $Filter -File
    if ($fileList.Count -eq 0) {
        return $null
    }

    # --- Prepare all ShellFile objects before timing ---
    Write-Host "Preparing $($fileList.Count) file objects..." -ForegroundColor Gray
    $shellObjects = foreach ($file in $fileList) {
        [Microsoft.WindowsAPICodePack.Shell.ShellFile]::FromFilePath($file.FullName)
    }

    # --- Start the timer and run the benchmark ---
    $stopwatch = [System.Diagnostics.Stopwatch]::new()
    $stopwatch.Start()

    foreach ($shellFile in $shellObjects) {
        try {
            # Set retrieval flags based on parameters
            if ($ThumbnailOnly) {
                $shellFile.Thumbnail.FormatOption = [Microsoft.WindowsAPICodePack.Shell.ShellThumbnailFormatOption]::ThumbnailOnly
            }
            if ($InCacheOnly) {
                $shellFile.Thumbnail.RetrievalOption = [Microsoft.WindowsAPICodePack.Shell.ShellThumbnailRetrievalOption]::CacheOnly
            }

            # Get the thumbnail
            $thumbnailBitmap = switch ($ThumbnailPreset) {
                "Small"      { $shellFile.Thumbnail.SmallBitmap }
                "Medium"     { $shellFile.Thumbnail.MediumBitmap }
                "Large"      { $shellFile.Thumbnail.LargeBitmap }
                "ExtraLarge" { $shellFile.Thumbnail.ExtraLargeBitmap }
                default      { $shellFile.Thumbnail.LargeBitmap }
            }
            $null = $thumbnailBitmap
        }
        catch {
            # Silently continue
        }
    }

    $stopwatch.Stop()

    # --- Create and return a result object ---
    $elapsed = $stopwatch.Elapsed
    $totalFiles = $shellObjects.Count
    $msPerFile = 0
    if ($totalFiles -gt 0) {
        $msPerFile = $elapsed.TotalMilliseconds / $totalFiles
    }

    return [PSCustomObject]@{
        TotalSeconds               = $elapsed.TotalSeconds
        AverageMillisecondsPerFile = $msPerFile
        FilesProcessed             = $totalFiles
        PresetUsed                 = $ThumbnailPreset
    }
}

Write-Host "`nPreparing to run benchmark $numberOfRuns times with preset '$size'..."
$allRunResults = @()

# --- Main Benchmark Loop ---
for ($i = 1; $i -le $numberOfRuns; $i++) {
    Write-Host "`n--- Starting Run $i of $numberOfRuns ---"
    $runResult = Benchmark-Thumbnails -Path $dir -ThumbnailPreset $size -ThumbnailOnly $thumbnailOnly -InCacheOnly $inCacheOnly -ApiPackDllPath $ApiPackDllPath

    if ($runResult) {
        $allRunResults += $runResult
        $totalTimeFormatted = "{0:N2}" -f $runResult.TotalSeconds
        $avgTimeFormatted = "{0:N2}" -f $runResult.AverageMillisecondsPerFile
        Write-Host "Run $i Complete. Total Time: $($totalTimeFormatted)s | Avg Time/File: $($avgTimeFormatted)ms"
    } else {
        Write-Warning "Run $i failed or found no files to process."
    }
}

# --- Calculate and Display Final Averages ---
if ($allRunResults.Count -gt 0) {
    $avgTotalTime = $allRunResults.TotalSeconds | Measure-Object -Average | Select-Object -ExpandProperty Average
    $avgOfAvgsPerFile = $allRunResults.AverageMillisecondsPerFile | Measure-Object -Average | Select-Object -ExpandProperty Average

    $avgTotalTimeFormatted = "{0:N2}" -f $avgTotalTime
    $avgOfAvgsFormatted = "{0:N2}" -f $avgOfAvgsPerFile

    Write-Host "`n----------------------------------" -ForegroundColor Green
    Write-Host "        OVERALL AVERAGE RESULTS (After $($allRunResults.Count) Runs)      " -ForegroundColor Green
    Write-Host "----------------------------------" -ForegroundColor Green
    Write-Host "Average Total Run Time:     $($avgTotalTimeFormatted) seconds"
    Write-Host "Average of Avg. Time/File:  $($avgOfAvgsFormatted) ms"
    Write-Host "----------------------------------" -ForegroundColor Green
} else {
    Write-Warning "`nBenchmark did not complete any runs successfully. No averages to display."
}


# --- Save first two thumbnails as samples for verification if desired ---
if ($outputSampleFiles) {
    $sampleFiles = Get-ChildItem -Path $dir -Filter '*.svg' -File | Select-Object -First 2
    foreach ($file in $sampleFiles) {
        try {
            $shellFile = [Microsoft.WindowsAPICodePack.Shell.ShellFile]::FromFilePath($file.FullName)

            # Select the correct property for the sample image
            $thumbnailBitmap = switch ($size) {
                "Small"      { $shellFile.Thumbnail.SmallBitmap }
                "Medium"     { $shellFile.Thumbnail.MediumBitmap }
                "Large"      { $shellFile.Thumbnail.LargeBitmap }
                "ExtraLarge" { $shellFile.Thumbnail.ExtraLargeBitmap }
                default      { $shellFile.Thumbnail.LargeBitmap }
            }

            if ($thumbnailBitmap) {
                # Save the sample next to the script file
                $outputFileName = "$($file.BaseName)_thumbnail_$($size)_sample.bmp"
                $outputPath = Join-Path -Path $PSScriptRoot -ChildPath $outputFileName

                $thumbnailBitmap.Save($outputPath, [System.Drawing.Imaging.ImageFormat]::Bmp)
                Write-Host " - Saved sample: $outputFileName"
            }
        }
        catch {
            Write-Warning "Could not generate sample for $($file.Name). Error: $_"
        }
    }
}
Read-Host -Prompt "Press Enter to exit"
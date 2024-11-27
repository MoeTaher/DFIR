# Define parameters
$BaseDir = "C:\DFIR"
$ThorZipUrl = "https://file.io/fJOccp9bkfOt"
$ThorZipPath = "$BaseDir\NN-Thor-ScannerNCA.zip"
$UnzipDir = "$BaseDir\NN-thor-scanner"
$NestedThorDir = "$UnzipDir\NN-Thor-Scanner"
$ThorExePath = "$UnzipDir\thor-lite.exe"
$JsonFilePath = "$BaseDir\thor_results.json"

$CommandLineArgs = @(
    "--utc",
    "--rfc3339",
    "--nocsv",
    "--nolog",
    "--nothordb",
    "--module", "Filescan",
    "--allhds",
    "--rebase-dir", $BaseDir,
    "--jsonfile", $JsonFilePath,
    "--json"
)

# Error handling function
function Handle-Error {
    param ([string]$Message)
    Write-Output "Error: $Message"
    exit 1
}

# Ensure any running thor-lite process is terminated
Write-Output "Checking for existing THOR (LITE) processes..."
try {
    $ThorProcesses = Get-Process -Name "thor-lite" -ErrorAction SilentlyContinue
    if ($ThorProcesses) {
        Write-Output "Terminating existing THOR (LITE) processes..."
        $ThorProcesses | Stop-Process -Force
        Write-Output "Processes terminated."
    } else {
        Write-Output "No existing THOR (LITE) processes found."
    }
} catch {
    Handle-Error "Failed to terminate existing THOR (LITE) processes."
}

# Remove the BaseDir
Write-Output "Attempting to remove folder: $BaseDir..."
try {
    if (Test-Path $BaseDir) {
        Remove-Item -Path $BaseDir -Recurse -Force
        Write-Output "Folder removed: $BaseDir"
    } else {
        Write-Output "Folder does not exist, skipping removal."
    }
} catch {
    Write-Output "Failed to remove folder: $BaseDir. Continuing execution."
}

# Ensure the BaseDir exists
Write-Output "Ensuring directory: $BaseDir"
try {
    if (-Not (Test-Path $BaseDir)) {
        New-Item -Path $BaseDir -ItemType Directory | Out-Null
        Write-Output "Directory created: $BaseDir"
    } else {
        Write-Output "Directory exists: $BaseDir"
    }
} catch {
    Handle-Error "Failed to create directory: $BaseDir"
}

# Add Defender exclusion for BaseDir
Write-Output "Adding Defender exclusion for: $BaseDir"
try {
    Add-MpPreference -ExclusionPath $BaseDir -Force
    Write-Output "Defender exclusion added for: $BaseDir"
} catch {
    Write-Output "Warning: Failed to add Defender exclusion for: $BaseDir. Continuing without exclusion."
}

# Check if the THOR ZIP file exists
if (-Not (Test-Path $ThorZipPath)) {
    Write-Output "Downloading THOR (LITE) binary from $ThorZipUrl to $ThorZipPath..."
    try {
        Invoke-WebRequest -Uri $ThorZipUrl -OutFile $ThorZipPath -UseBasicParsing
        Write-Output "Download completed: $ThorZipPath"
    } catch {
        Handle-Error "Failed to download THOR (LITE) binary."
    }
} else {
    Write-Output "THOR (LITE) binary already exists at $ThorZipPath. Skipping download."
}

# Verify the ZIP file
Write-Output "Verifying the ZIP file: $ThorZipPath..."
try {
    if (-Not (Test-Path $ThorZipPath)) {
        Handle-Error "ZIP file does not exist: $ThorZipPath"
    }
    $FileSize = (Get-Item $ThorZipPath).Length
    if ($FileSize -lt 1024) {
        Handle-Error "ZIP file is too small to be valid. Size: $FileSize bytes"
    }
    Write-Output "ZIP file verification passed."
} catch {
    Handle-Error "Failed to verify the ZIP file: $ThorZipPath"
}

# Unzip the binary into UnzipDir
Write-Output "Unzipping THOR (LITE) binary to $UnzipDir..."
try {
    if (Test-Path $UnzipDir) {
        Remove-Item -Path $UnzipDir -Recurse -Force
    }
    New-Item -Path $UnzipDir -ItemType Directory | Out-Null
    Expand-Archive -Path $ThorZipPath -DestinationPath $UnzipDir -Force

    # Move files from nested directory if it exists
    if (Test-Path $NestedThorDir) {
        Get-ChildItem -Path $NestedThorDir -Recurse | Move-Item -Destination $UnzipDir -Force
        Remove-Item -Path $NestedThorDir -Recurse -Force
    }
    Write-Output "Unzip and reorganization completed."
} catch {
    Write-Output "Unzip using Expand-Archive failed. Attempting alternative method..."
    # Fallback to extracting using COM object if Expand-Archive fails
    try {
        $Shell = New-Object -ComObject Shell.Application
        $Zip = $Shell.NameSpace($ThorZipPath)
        if ($Zip) {
            $Shell.NameSpace($UnzipDir).CopyHere($Zip.Items(), 4)
            Start-Sleep -Seconds 5  # Allow time for files to copy
            Write-Output "Fallback unzip completed."
        } else {
            Handle-Error "Failed to extract using fallback method."
        }
    } catch {
        Handle-Error "Both unzip methods failed."
    }
}

# Verify the executable exists
Write-Output "Verifying THOR (LITE) executable at $ThorExePath..."
if (-Not (Test-Path $ThorExePath)) {
    Handle-Error "THOR (LITE) executable not found at $ThorExePath."
}

# Run THOR (LITE) with the provided command line arguments
Write-Output "Executing THOR (LITE)..."
try {
    Start-Process -FilePath $ThorExePath -ArgumentList $CommandLineArgs -Wait -NoNewWindow
    Write-Output "THOR (LITE) execution completed."
} catch {
    Handle-Error "Failed to execute THOR (LITE)."
}

Write-Output "THOR (LITE) runner completed successfully."

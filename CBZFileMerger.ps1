# Prompt the user for the folder path

Function Main {
$folderPath = Read-Host "Enter the folder path where your manga CBZ/ZIP files are located (Press Enter to use default path)"

# Set a default folder path if the user doesn't provide one
if ([string]::IsNullOrWhiteSpace($folderPath)) {
    $folderPath = "C:\Users\bhaktaravin\Documents\Mangas"
    Write-Host "Using default folder path: $folderPath"
}

# Validate that the folder exists
if (-Not (Test-Path $folderPath)) {
    Write-Error "The folder '$folderPath' does not exist or is not a valid path."
    return
}

# Get all .cbz and .zip files in the folder
$cbzFiles = Get-ChildItem -Path $folderPath -Recurse -Include "*.cbz", "*.zip" -File

# Check if there are any CBZ/ZIP files
if ($cbzFiles.Count -eq 0) {
    Write-Error "No .cbz or .zip files found in the folder '$folderPath'."
    return
}

# Create a temporary folder for extracted contents
$extractFolder = Join-Path -Path $folderPath -ChildPath "ExtractedContents"
if (-Not (Test-Path $extractFolder)) {
    New-Item -ItemType Directory -Path $extractFolder | Out-Null
}

# Extract contents of each CBZ/ZIP file
Write-Host "Extracting contents of CBZ/ZIP files..."
foreach ($cbzFile in $cbzFiles) {
    Write-Host "Processing file: $($cbzFile.FullName)"

    # Check if the file is locked
    if (Test-FileLock $cbzFile.FullName) {
        Write-Host "The file $($cbzFile.FullName) is locked by another process. Skipping."
        continue
    }

    # Rename .cbz to .zip if necessary
    $zipFilePath = $cbzFile.FullName
    if ($cbzFile.Extension -eq ".cbz") {
        $zipFilePath = $cbzFile.FullName -replace '\.cbz$', '.zip'
        Rename-Item -Path $cbzFile.FullName -NewName $zipFilePath -ErrorAction Stop
    }

    # Extract the ZIP file
    try {
        $destination = Join-Path -Path $extractFolder -ChildPath ($cbzFile.BaseName)
        Expand-Archive -Path $zipFilePath -DestinationPath $destination -Force
        Write-Host "Extracted: $zipFilePath to $destination"
    } catch {
        Write-Error "Failed to extract '$zipFilePath'. Error: $_"
    }
}

# Merge all extracted contents into a single folder
$mergedFolder = Join-Path -Path $folderPath -ChildPath "MergedContents"
if (-Not (Test-Path $mergedFolder)) {
    New-Item -ItemType Directory -Path $mergedFolder | Out-Null
}

Write-Host "Merging contents..."
Get-ChildItem -Path $extractFolder -Recurse | ForEach-Object {
    if ($_.PSIsContainer -eq $false) {
        $destinationPath = Join-Path -Path $mergedFolder -ChildPath $_.Name
        Copy-Item -Path $_.FullName -Destination $destinationPath -Force
    }
}

# Create a new CBZ file named after the folder
$folderName = Split-Path -Path $folderPath -Leaf
$mergedZIPPath = Join-Path -Path $folderPath -ChildPath "$folderName.zip"
$mergedCBZPath = Join-Path -Path $folderPath -ChildPath "$folderName.cbz"

Write-Host "Creating merged CBZ file..."
if (Test-Path $mergedZIPPath) {
    Remove-Item -Path $mergedZIPPath -Force
}

Compress-Archive -Path $mergedFolder\* -DestinationPath $mergedZIPPath -Force

# Rename the ZIP file to CBZ
if (Test-Path $mergedCBZPath) {
    Remove-Item -Path $mergedCBZPath -Force
}

Rename-Item -Path $mergedZIPPath -NewName $mergedCBZPath -Force

Write-Host "Merged CBZ file created at: $mergedCBZPath"

# Cleanup temporary files
Write-Host "Cleaning up temporary files..."
Remove-Item -Path $extractFolder -Recurse -Force
Remove-Item -Path $mergedFolder -Recurse -Force

Write-Host "Process complete!"
}
# Function to check if a file is locked
# Function to check if a file is locked by another process
Function Test-FileLock {
    param(
        [string]$FilePath
    )
    try {
        # Attempt to open the file in ReadWrite mode
        $stream = [System.IO.File]::Open($FilePath, 'Open', 'ReadWrite', 'None')
        $stream.Close()
        return $false  # File is not locked
    } catch {
        return $true   # File is locked
    }
}

Main
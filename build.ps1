# PowerShell script to build the executable using PyInstaller

# Install required dependencies if not already installed
pip install pywin32 keyboard pystray Pillow PyInstaller PyQt6

# Get the directory of the PowerShell script (same as the .py file)
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Run PyInstaller to build the executable
pyinstaller --onefile --windowed seethru.py
#pyinstaller --onefile --windowed --icon=icon.ico seethru.py

# Move the .exe from dist to the script directory
$exeSource = Join-Path -Path $scriptDir -ChildPath "dist\seethru.exe"
$exeDestination = Join-Path -Path $scriptDir -ChildPath "seethru.exe"
if (Test-Path $exeSource) {
    Move-Item -Path $exeSource -Destination $exeDestination -Force
    Write-Host "Moved executable to $exeDestination"
} else {
    Write-Host "Error: Executable not found in dist folder."
    exit 1
}

# Delete the dist and build folders
$distFolder = Join-Path -Path $scriptDir -ChildPath "dist"
$buildFolder = Join-Path -Path $scriptDir -ChildPath "build"
if (Test-Path $distFolder) {
    Remove-Item -Path $distFolder -Recurse -Force
    Write-Host "Deleted dist folder"
}
if (Test-Path $buildFolder) {
    Remove-Item -Path $buildFolder -Recurse -Force
    Write-Host "Deleted build folder"
}

Remove-Item .\seethru.spec

# Output message
Write-Host "Build complete. The executable is in $exeDestination"
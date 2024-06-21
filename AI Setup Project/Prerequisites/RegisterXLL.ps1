param (
    [string]$installationPath
)

# Determine if the process is 64-bit
$is64BitProcess = $env:PROCESSOR_ARCHITECTURE -eq "AMD64"

# Set the add-in paths
$xll32Path = Join-Path -Path $installationPath -ChildPath "KDC Excel-AddIn.xll"
$xll64Path = Join-Path -Path $installationPath -ChildPath "KDC Excel-AddIn64.xll"

# Check if either file exists
if (-Not (Test-Path -Path $xll32Path) -and -Not (Test-Path -Path $xll64Path)) {
    Write-Error "Error: Neither add-in file could be found."
    exit 1
}

# Choose the correct add-in path based on the bitness
$addInPath = if ($is64BitProcess -and (Test-Path -Path $xll64Path)) {
    $xll64Path
} else {
    $xll32Path
}

# Log the path for debugging
Write-Output "Registering XLL: $addInPath"

# Function to load the XLL
function Register-XLL($path) {
    $excel = New-Object -ComObject Excel.Application
    $excel.Workbooks.Add() | Out-Null
    $excel.RegisterXLL($path) | Out-Null
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

# Call the function to register the XLL
Register-XLL -path $addInPath

$packageName = 'rebis-grade-excel-importer'
$url = 'https://github.com/ucoruh/rebis-grade-excel-importer/releases/download/v1.0.0/rebis-grade-excel-importer.zip'
$installPath = "$env:ProgramFiles\Rebis\GradeExcelImporter"

# Download the necessary files
$zipFilePath = Join-Path $PSScriptRoot "$packageName.zip"
Invoke-WebRequest -Uri $url -OutFile $zipFilePath

# Install any required dependencies or tools (none required for this app)

# Extract the application files to the install location
New-Item -ItemType Directory -Path $installPath -Force
Expand-Archive -Path $zipFilePath -DestinationPath $installPath -Force

# Clean up the zip file
Remove-Item $zipFilePath -Force
